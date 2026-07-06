import json
import re
from pathlib import Path

# Paths
SCRIPT_DIR = Path(__file__).resolve().parent
WORKSPACE_ROOT = SCRIPT_DIR.parent
DEFAULT_CONFIG_PATH = WORKSPACE_ROOT / "config" / "comparable_assembly_rules.json"

def load_rules(config_path=DEFAULT_CONFIG_PATH):
    """Loads the comparable assembly rules from the JSON configuration file."""
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        raise RuntimeError(f"Failed to load comparable assembly rules from {config_path}: {e}")

def normalize(val):
    """Normalizes input string for comparison by lowercasing, stripping,

    collapsing whitespaces, and resolving synonyms.
    """
    if val is None:
        return ""
    s = str(val).lower().strip()
    s = re.sub(r'\s+', ' ', s)
    
    # Resolve spelling variations and synonyms
    s = re.sub(r'\bma\b', 'mechanically attached', s)
    s = s.replace('dens deck', 'densdeck')
    s = s.replace('gypsum board', 'gypsum')
    s = s.replace('fanfold', 'fan fold')
    s = s.replace('rhino bond', 'rhinobond')
    s = s.replace('high density polyiso', 'hd polyiso')
    s = s.replace('hd iso', 'hd polyiso')
    s = s.replace('induction welded', 'rhinobond')
    return s

def find_group(val, rules):
    """Finds the group definition and matched term for a given assembly string (single-dimension logic)."""
    normalized_val = normalize(val)
    if not normalized_val or normalized_val in ["nan", "none", "n/a", "all"]:
        return None, None
        
    groups = rules.get("groups", {})
    
    # 1. Exact match against group terms
    for group_id, group_info in groups.items():
        for term in group_info.get("terms", []):
            if normalized_val == normalize(term):
                return group_id, group_info
                
    # 2. Substring/Word boundary match
    terms_to_check = []
    for group_id, group_info in groups.items():
        for term in group_info.get("terms", []):
            terms_to_check.append((normalize(term), group_id, group_info))
            
    terms_to_check.sort(key=lambda x: len(x[0]), reverse=True)
    
    for norm_term, group_id, group_info in terms_to_check:
        if len(norm_term) <= 2:
            pattern = r'\b' + re.escape(norm_term) + r'\b'
            if re.search(pattern, normalized_val):
                return group_id, group_info
        else:
            if norm_term in normalized_val:
                return group_id, group_info
                
    return None, None

def compare_assemblies(val1, val2, rules=None):
    """Compares two roofing assembly strings (single-dimension fallback for backward compatibility)."""
    if rules is None:
        rules = load_rules()
        
    norm1 = normalize(val1)
    norm2 = normalize(val2)
    
    # 1. Handle exact matches
    if norm1 == norm2 and norm1 not in ["", "nan", "none", "n/a", "all"]:
        group_id, group_info = find_group(val1, rules)
        match_group = group_info["name"] if group_info else "Exact Match"
        return 1, 1.00, "Exact match", match_group

    # 2. Identify groups
    group1_id, group1_info = find_group(val1, rules)
    group2_id, group2_info = find_group(val2, rules)
    
    if group1_id is None or group2_id is None:
        if norm1 == norm2:
            return 1, 1.00, "Exact match", "Exact Match"
        return None, 0.00, f"Incompatible assemblies: '{val1}' vs '{val2}'", None

    # 3. Check intra-group matching
    if group1_id == group2_id:
        return (
            group1_info["intra_group_tier"],
            group1_info["intra_group_weight"],
            group1_info["intra_group_reason"],
            group1_info["name"]
        )

    # 4. Check inter-group relationships
    relationships = rules.get("inter_group_relationships", [])
    for rel in relationships:
        g1, g2 = rel["group1"], rel["group2"]
        if (group1_id == g1 and group2_id == g2) or (group1_id == g2 and group2_id == g1):
            combined_name = f"{group1_info['name']} / {group2_info['name']}"
            return (
                rel["tier"],
                rel["weight"],
                rel["reason"],
                combined_name
            )
            
    return None, 0.00, f"No relationship defined between groups: '{group1_info['name']}' and '{group2_info['name']}'", None

# =====================================================================
# Phase 2 Multi-Dimensional Matching Engine
# =====================================================================

def extract_assembly_dimensions(text, rules=None):
    """Parses and extracts matching terms for the three distinct dimensions

    (membrane, board, attachment) from a single assembly string.
    """
    if rules is None:
        rules = load_rules()
        
    norm_text = normalize(text)
    groups = rules.get("groups", {})
    
    dimensions = {
        "membrane": {
            "detected": False,
            "group_id": None,
            "group_name": None,
            "matched_terms": []
        },
        "board": {
            "detected": False,
            "group_id": None,
            "group_name": None,
            "matched_terms": []
        },
        "attachment": {
            "detected": False,
            "group_id": None,
            "group_name": None,
            "matched_terms": []
        }
    }
    
    if not norm_text or norm_text in ["nan", "none", "n/a", "all", "unknown coating system"]:
        return dimensions
        
    for dim in ["membrane", "board", "attachment"]:
        matches = []
        for group_id, group_info in groups.items():
            if group_info.get("dimension") != dim:
                continue
            for term in group_info.get("terms", []):
                norm_term = normalize(term)
                matched = False
                if len(norm_term) <= 2:
                    pattern = r'\b' + re.escape(norm_term) + r'\b'
                    if re.search(pattern, norm_text):
                        matched = True
                else:
                    if norm_term in norm_text:
                        matched = True
                        
                if matched:
                    matches.append((norm_term, group_id, group_info))
                    
        if matches:
            # Sort by matched term length descending to pick the most specific match
            matches.sort(key=lambda x: len(x[0]), reverse=True)
            best_match_term, best_match_group_id, best_match_group_info = matches[0]
            
            # Extract all matched terms for this group in the text
            group_matched_terms = []
            for term in best_match_group_info.get("terms", []):
                norm_term = normalize(term)
                if len(norm_term) <= 2:
                    pattern = r'\b' + re.escape(norm_term) + r'\b'
                    if re.search(pattern, norm_text):
                        group_matched_terms.append(term)
                else:
                    if norm_term in norm_text:
                        group_matched_terms.append(term)
                        
            # Unique matched terms preserving order
            unique_matched_terms = list(dict.fromkeys(group_matched_terms))
            
            dimensions[dim] = {
                "detected": True,
                "group_id": best_match_group_id,
                "group_name": best_match_group_info["name"],
                "matched_terms": unique_matched_terms
            }
            
    return dimensions

def compare_assembly_dimensions(text1, text2, rules=None):
    """Compares each dimension (membrane, board, attachment) independently,

    redistributing weight when a dimension is missing in both descriptions.
    """
    if rules is None:
        rules = load_rules()
        
    d1 = extract_assembly_dimensions(text1, rules)
    d2 = extract_assembly_dimensions(text2, rules)
    
    std_weights = {
        "membrane": 0.35,
        "board": 0.40,
        "attachment": 0.25
    }
    
    # 1. Determine active dimensions (present in at least one)
    active_dimensions = []
    for dim in ["membrane", "board", "attachment"]:
        if d1[dim]["detected"] or d2[dim]["detected"]:
            active_dimensions.append(dim)
            
    # 2. Handle case where no dimensions are detected
    if not active_dimensions:
        return {
            "overall_tier": "Excluded / Not Comparable",
            "overall_weight": 0.00,
            "overall_reason": "No matching dimensions found in either description",
            "dimension_results": {
                "membrane": {"status": "missing_both", "score": 0.00, "reason": "Not detected in either description"},
                "board": {"status": "missing_both", "score": 0.00, "reason": "Not detected in either description"},
                "attachment": {"status": "missing_both", "score": 0.00, "reason": "Not detected in either description"}
            }
        }
        
    # 3. Calculate redistributed weights
    sum_active_weights = sum(std_weights[dim] for dim in active_dimensions)
    redist_weights = {}
    for dim in ["membrane", "board", "attachment"]:
        if dim in active_dimensions:
            redist_weights[dim] = std_weights[dim] / sum_active_weights
        else:
            redist_weights[dim] = 0.00
            
    # 4. Process each dimension
    dim_results = {}
    overall_score = 0.00
    reasons_summary = []
    
    for dim in ["membrane", "board", "attachment"]:
        info1 = d1[dim]
        info2 = d2[dim]
        
        # If missing in both
        if not info1["detected"] and not info2["detected"]:
            dim_results[dim] = {
                "status": "ignored",
                "score": 0.00,
                "reason": "Missing in both descriptions (weight redistributed)"
            }
            continue
            
        # If present in one but missing in the other
        if info1["detected"] != info2["detected"]:
            who_has = "first" if info1["detected"] else "second"
            warn_msg = f"{dim.capitalize()} present in {who_has} description but missing in the other"
            dim_results[dim] = {
                "status": "mismatch_missing",
                "score": 0.00,
                "reason": f"Warning: {warn_msg}"
            }
            reasons_summary.append(f"{dim.capitalize()}: mismatch (present in one only, scored 0.00)")
            continue
            
        # Present in both, compare groups and terms
        g1_id = info1["group_id"]
        g2_id = info2["group_id"]
        
        # Check for exact term intersection (based on normalized sets)
        terms1_norm = {normalize(t) for t in info1["matched_terms"]}
        terms2_norm = {normalize(t) for t in info2["matched_terms"]}
        
        if g1_id == g2_id:
            # Same group
            if terms1_norm == terms2_norm:
                dim_score = 1.00
                dim_tier = 1
                dim_reason = "Exact match"
            else:
                group_info = rules["groups"][g1_id]
                dim_score = group_info["intra_group_weight"]
                dim_tier = group_info["intra_group_tier"]
                dim_reason = group_info["intra_group_reason"]
        else:
            # Different groups, check inter-group relations
            relationships = rules.get("inter_group_relationships", [])
            found_rel = None
            for rel in relationships:
                if (g1_id == rel["group1"] and g2_id == rel["group2"]) or (g1_id == rel["group2"] and g2_id == rel["group1"]):
                    found_rel = rel
                    break
                    
            if found_rel:
                dim_score = found_rel["weight"]
                dim_tier = found_rel["tier"]
                dim_reason = found_rel["reason"]
            else:
                dim_score = 0.00
                dim_tier = None
                dim_reason = f"No relationship defined between group '{info1['group_name']}' and '{info2['group_name']}'"
                
        dim_results[dim] = {
            "status": "matched",
            "score": dim_score,
            "tier": dim_tier,
            "reason": dim_reason,
            "group1": info1["group_name"],
            "group2": info2["group_name"],
            "terms1": info1["matched_terms"],
            "terms2": info2["matched_terms"]
        }
        
        overall_score += redist_weights[dim] * dim_score
        reasons_summary.append(f"{dim.capitalize()}: matched ({info1['group_name']} vs {info2['group_name']}, Tier {dim_tier}, score {dim_score:.2f})")
        
    # 5. Determine overall tier mapping
    if overall_score >= 0.90:
        overall_tier = "Tier 1, Direct Equivalent"
    elif overall_score >= 0.60:
        overall_tier = "Tier 2, Related Assembly"
    elif overall_score >= 0.30:
        overall_tier = "Tier 3, Directional Benchmark"
    else:
        overall_tier = "Excluded / Not Comparable"
        
    # Append weight redistribution details if applicable
    redist_details = []
    for dim in ["membrane", "board", "attachment"]:
        if dim in active_dimensions and abs(redist_weights[dim] - std_weights[dim]) > 1e-5:
            redist_details.append(f"{dim}: {std_weights[dim]:.2f}->{redist_weights[dim]:.2f}")
    if redist_details:
        reasons_summary.append(f"Weight redistribution: {', '.join(redist_details)}")
        
    overall_reason = "; ".join(reasons_summary)
    
    return {
        "overall_tier": overall_tier,
        "overall_weight": round(overall_score, 4),
        "overall_reason": overall_reason,
        "dimension_results": dim_results
    }
