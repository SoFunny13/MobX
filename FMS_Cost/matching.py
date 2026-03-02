import learned_mappings

PLATFORM_SUFFIXES = [" Android", " iOS", " AOS", " Web", " Мобайл", " android", " ios"]

# Aliases: tracker name (from stats) -> canonical name (in reference file)
SOURCE_ALIASES = {
    "xiaomiglobal_int": "xiaomi",
    "xiaomiglobal_int_contract": "xiaomi",
    "yandexdirect_int": "Яндекс Директ",
}


def normalize(name):
    return " ".join(name.lower().strip().split())


def normalize_nospace(name):
    return name.lower().strip().replace(" ", "")


def strip_platform_suffix(offer_name):
    for suffix in PLATFORM_SUFFIXES:
        if offer_name.endswith(suffix):
            return offer_name[: -len(suffix)]
    return offer_name


def match_offer(stats_name, offers_dict):
    """Match offer name from stats file against reference dict.

    offers_dict: normalized_name -> (id, original_name)
    Returns offer_id (int) or None.
    """
    # Check learned mappings first
    learned_id = learned_mappings.get_offer_id(stats_name.strip())
    if learned_id is not None:
        return learned_id

    base = strip_platform_suffix(stats_name.strip())
    key = normalize(base)

    # Exact normalized match
    if key in offers_dict:
        return offers_dict[key][0]

    # No-space match (catches "Быстроденьги" vs "Быстро Деньги")
    key_nospace = normalize_nospace(base)
    for ref_name, (ref_id, _) in offers_dict.items():
        if normalize_nospace(ref_name) == key_nospace:
            return ref_id

    # Substring containment
    for ref_name, (ref_id, _) in offers_dict.items():
        if key in ref_name or ref_name in key:
            return ref_id

    return None


def match_source(stats_name, sources_dict):
    """Match source name from stats file against reference dict.

    sources_dict: normalized_name -> (id, original_name)
    Returns source_id (int) or None.
    """
    raw = stats_name.strip()

    # Check learned mappings first
    learned_id = learned_mappings.get_source_id(raw)
    if learned_id is not None:
        return learned_id

    # Check aliases
    alias = SOURCE_ALIASES.get(raw.lower())
    if alias:
        alias_key = normalize(alias)
        if alias_key in sources_dict:
            return sources_dict[alias_key][0]

    key = normalize(raw)

    # Exact normalized match
    if key in sources_dict:
        return sources_dict[key][0]

    # Strip _int / _int_contract suffix and try matching
    base = stats_name.lower().replace("_int_contract", "").replace("_int", "").strip()
    for ref_name, (ref_id, _) in sources_dict.items():
        ref_lower = ref_name.lower()
        if base == ref_lower:
            return ref_id
        if base in ref_lower.replace(" ", ""):
            return ref_id

    return None
