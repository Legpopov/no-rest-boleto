"""Explicit, geographically bounded sector proxy; never an official forecast."""
from __future__ import annotations

from .planning import positive


def category_index(root, category):
    row = root['indices'][category]
    if not positive(row['index']) or not root.get('source'):
        raise ValueError('Invalid sourced category index')
    if root.get('unit') != 'percent_of_same_month_previous_year':
        raise ValueError('Index must be same-month YoY, not month-on-month or YTD')
    return dict(rate=row['index']/100-1,source=root['source'],
                reference_period=root['reference_period'],comparison_period=root['comparison_period'],
                method='official_category_yoy_repeated_scenario',scope=row['scope'])


def domestic_pair(origin, destination, registry, root):
    excluded = set(root['excluded_territories'])
    return (origin not in excluded and destination not in excluded
            and registry[origin].get('country_code') == 'RU'
            and registry[destination].get('country_code') == 'RU')
