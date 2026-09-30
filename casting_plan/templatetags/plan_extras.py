from django import template

register = template.Library()

@register.filter
def get_item(dictionary, key):
    """
    Safely retrieve a dictionary value by key, supporting exact match,
    case-insensitive match, and normalized key match (spaces/underscores/dots).
    """
    if not isinstance(dictionary, dict):
        return ""
    if key in dictionary and dictionary[key] is not None:
        return dictionary[key]
    
    key_str = str(key).strip()
    # Direct match
    for k, v in dictionary.items():
        if str(k).strip() == key_str:
            return v if v is not None else ""

    # Case-insensitive match
    key_lower = key_str.lower()
    for k, v in dictionary.items():
        if str(k).strip().lower() == key_lower:
            return v if v is not None else ""

    # Normalized match (replace _ with space and remove .)
    norm_key = key_lower.replace('_', ' ').replace('.', '').replace('-', '')
    for k, v in dictionary.items():
        norm_k = str(k).lower().replace('_', ' ').replace('.', '').replace('-', '')
        if norm_k == norm_key:
            return v if v is not None else ""
            
    return ""
