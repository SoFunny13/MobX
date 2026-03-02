import json
import os

import config

MAPPINGS_FILE = os.path.join(config.BASE_DIR, "learned_mappings.json")

_cache = None


def _load():
    global _cache
    if os.path.exists(MAPPINGS_FILE):
        with open(MAPPINGS_FILE, "r", encoding="utf-8") as f:
            _cache = json.load(f)
    else:
        _cache = {"offers": {}, "sources": {}}
    return _cache


def _save():
    with open(MAPPINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(_cache, f, ensure_ascii=False, indent=2)


def get():
    return _load()


def save_offers(mappings):
    """Save new offer name -> id mappings. mappings: dict {name: id}."""
    data = get()
    data["offers"].update(mappings)
    _save()


def save_sources(mappings):
    """Save new source name -> id mappings. mappings: dict {name: id}."""
    data = get()
    data["sources"].update(mappings)
    _save()


def get_offer_id(name):
    """Look up offer ID from learned mappings. Returns int or None."""
    val = get()["offers"].get(name)
    return int(val) if val is not None else None


def get_source_id(name):
    """Look up source ID from learned mappings. Returns int or None."""
    val = get()["sources"].get(name)
    return int(val) if val is not None else None
