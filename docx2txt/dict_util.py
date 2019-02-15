"""Dictionary Utilities"""


def merge(dicts):
    # type: (list) -> dict
    merged = {}  # type: dict
    for d in dicts:
        merged.update(d)

    return merged


def filter_key(src_dict, key_test, getter=dict.values):
    return getter({
        key: val
        for key, val
        in src_dict.items()
        if key_test(key)})
