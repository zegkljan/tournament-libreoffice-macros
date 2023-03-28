# coding: utf-8

import math
from typing import List, Optional, Tuple, Union
import typing

T = typing.TypeVar('T')


def findGroupSizes(n_total: int, max_group_size: int) -> List[int]:
    """Based on the total number of participants and maximum allowed size of a group,
    determines sizes of individual groups.
    """
    rem = n_total % max_group_size
    n = n_total // max_group_size
    if rem == 0:
        return [max_group_size] * n
    elif max_group_size <= 4:
        raise ValueError('Cannot arrange groups - largest group would be smaller than 4.')
    elif n + 1 >= max_group_size - rem:
        groups = [max_group_size] * (n + 1 - (max_group_size - rem))
        groups += [max_group_size - 1] * (max_group_size - rem)
        groups.sort()
        return groups
    else:
        return findGroupSizes(n_total, max_group_size - 1)


def assignGroups(group_sizes: List[int], participants: List[T]) -> List[List[T]]:
    """Assigns participants into groups of the given sizes.

    Uses the 'snake' algorithm, assuming that participants are already sorted by the desired ranking.
    """
    min_size = min(group_sizes)
    groups = [[] for _ in range(len(group_sizes))]
    assigned = 0
    g = 0
    d = 0
    i = 0
    while i < len(participants):
        p = participants[i]
        if group_sizes[g] > len(groups[g]):
            groups[g].append(p)
            assigned += 1
            i += 1
        if g == 0 or g == len(groups) - 1:
            if g != 0:
                d -= 1
            elif g != len(groups) - 1:
                d += 1
        g += d
    if assigned != len(participants):
        raise ValueError('not all participants assigned - this should not happen')
    return groups


def makeGroupSchedule(group: List[T]):
    """Given a group, returns a list of pairs representing the individual matches in the group.

    For groups of even-numbered size, uses the 'circle' algorithm, and for groups of odd-numbered size uses the algorithm from https://arxiv.org/abs/1804.04504v1.
    """
    if len(group) % 2 == 0:
        return makeGroupCircle(group)
    else:
        return makeGroupOdd(group)


def makeGroupCircle(group: List[T]) -> List[Tuple[T, T]]:
    """Schedules matches in group according to the 'circle' algorithm.
    """
    if len(group) % 2 == 1:
        group = [None] + group
    schedule = []
    circle = list(group)
    schedule.extend([(circle[i], circle[len(circle) - i - 1])
                     for i in range(len(group) // 2)
                     if circle[i] is not None and circle[len(circle) - i - 1] is not None
                    ])
    for k in range(len(group) - 2):
        tmp = circle[-1]
        circle[2:] = circle[1:-1]
        circle[1] = tmp
        if k % 2 == 0:
            circle[0], circle[-1] = circle[-1], circle[0]
        schedule.extend([(circle[i], circle[len(circle) - i - 1])
                         for i in range(len(group) // 2)
                         if circle[i] is not None and circle[len(circle) - i - 1] is not None
                        ])
        if k % 2 == 0:
            circle[0], circle[-1] = circle[-1], circle[0]
            
    return schedule


def makeGroupOdd(group):
    """Schedules matches in a group of odd-numbered size according to the algorithm from https://arxiv.org/abs/1804.04504v1.
    """
    assert len(group) % 2 == 1
    n = len(group)
    k = n // 2
    rounds = [[[None, None] for _ in range(k)] for _ in range(n)]

    def putToSlot(r, s, t, flip):
        slt = rounds[r - 1][s - 1]
        first = 1 if flip else 0
        second = 0 if flip else 1
        if slt[first] is None:
            slt[first] = group[t - 1]
        else:
            slt[second] = group[t - 1]

    for i in range(1, k + 1):
        team = 2 * i - 1
        slot = i
        flip = False
        for rnd in range(1, 2 * i + 1):
            putToSlot(rnd, slot, team, flip)
            flip = not flip
        for rnd in range(2 * i + 1, n + 1):
            slot = (slot + 1) % (k + 1)
            if slot == 0:
                continue
            putToSlot(rnd, slot, team, flip)
            flip = not flip

        team = 2 * i
        slot = i
        putToSlot(1, slot, team, flip)
        flip = True
        for rnd in range(2, 2 * k + 3 - 2 * i + 1):
            slot = (slot + 1) % (k + 1)
            if slot == 0:
                continue
            putToSlot(rnd, slot, team, flip)
            flip = not flip
        for rnd in range(2 * k + 3 - 2 * i + 1, n + 1):
            if slot == 0:
                continue
            putToSlot(rnd, slot, team, flip)
            flip = not flip
        
    team = 2 * k + 1
    flip = False
    for j in range(1, n + 1):
        slot = j // 2
        if slot == 0:
            continue
        putToSlot(j, slot, team, flip)
        flip = not flip
    
    group = [(a, b) for rnd in rounds for a, b in rnd]
    nums_first = dict()
    for a, _ in group:
        if a not in nums_first:
            nums_first[a] = 0
        nums_first[a] += 1
    while True:
        most = max(nums_first.items(), key=lambda x: x[1])[0]
        least = min(nums_first.items(), key=lambda x: x[1])[0]
        if nums_first[most] - nums_first[least] > 1:
            for i, (a, b) in enumerate(group):
                if a == most and b == least:
                    group[i] = (b, a)
                    nums_first[a] -= 1
                    nums_first[b] += 1
                    break
        else:
            break
    
    return group


def makeElimination(participants: List[T]) -> List[Tuple[Optional[T], Optional[T]]]:
    """Schedules 1st level of an elimination bracket.
    
    If the number of participants is not a power of 2, the 'extra' participants will be paired with None.
    It is assumed that the participants are already sorted by the desired ranking.
    """
    n = len(participants)
    n2log = math.ceil(math.log2(n))

    assert n > 2

    layer = [(0, 1)]
    for lvl in range(1, n2log):
        max_n = 2 ** (lvl + 1) - 1

        layer2 = []
        for a, b in layer:
            layer2.append((a, max_n - a))
            layer2.append((max_n - b, b))
        layer = layer2
    
    participants = participants + [None] * (2 ** n2log - n)
    res = [(participants[a], participants[b]) for a, b in layer]
    return res, n2log
