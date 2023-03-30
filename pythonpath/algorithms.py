# coding: utf-8

import math
from typing import List, Optional, Tuple, Union, Sequence, Callable, Any
import typing
from scipy.sparse import csr_matrix
from scipy.sparse.csgraph import maximum_bipartite_matching

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


def assignGroups(group_sizes: List[int], participants: List[T], spreadCriteriaGetters: List[Callable[[T], Any]]) -> List[List[T]]:
    """Assigns participants into groups of the given sizes.

    Tries to not put people who have the same value of spread criteria into the same groups.
    Where not possible, tends towards the 'snake' algorithm.
    Assumes that participants are already sorted by the desired ranking.
    """
    groups = [[] for _ in range(len(group_sizes))]

    def getLeastFilled():
        least = []
        for i, g in enumerate(groups):
            if len(g) == group_sizes[i]:
                continue
            if not least or len(g) < len(least[0]):
                least = [g]
            elif len(g) == len(least[0]):
                least.append(g)
        return least
    
    def assign(grps, parts, nCriteria):
        assert len(grps) == len(parts)
        m = csr_matrix([[1] * len(parts)] * len(grps))
        for ig, g in enumerate(grps):
            for ip, p in enumerate(parts):
                for gm in g:
                    if any((c(gm) == c(p) for c in spreadCriteriaGetters[:nCriteria])):
                        m[ig, ip] = 0
                        break
        m.eliminate_zeros()
        res = maximum_bipartite_matching(m, perm_type='row')
        return list(res)

    layer = 0
    participants = list(participants)
    while participants:
        least = getLeastFilled()
        if layer % 2 == 1:
            least = list(reversed(least))
        ps, participants = participants[:len(least)], participants[len(least):]
        assignment = None
        for critCounter in range(len(spreadCriteriaGetters) + 1):
            assignment = assign(least, ps, len(spreadCriteriaGetters) - critCounter)
            if all((x >= 0 for x in assignment)):
                break
        for i, n in enumerate(assignment):
            least[n].append(ps[i])
        layer += 1
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
    round = [(circle[i], circle[len(circle) - i - 1]) 
             for i in range(len(group) // 2)
             if circle[i] is not None and circle[len(circle) - i - 1] is not None
            ]
    schedule.extend(reversed(round))
    for k in range(len(group) - 2):
        circle = [circle[0], circle[-1]] + circle[1:-1]
        if k % 2 == 0:
            circle[0], circle[-1] = circle[-1], circle[0]
        round = [(circle[i], circle[len(circle) - i - 1])
                 for i in range(len(group) // 2)
                 if circle[i] is not None and circle[len(circle) - i - 1] is not None
                ]
        schedule.extend(reversed(round))
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
        if nums_first[most] - nums_first[least] > 0:
            for i, (a, b) in enumerate(group):
                if a == most and b == least:
                    group[i] = (b, a)
                    nums_first[a] -= 1
                    nums_first[b] += 1
                    break
        else:
            break
    
    return reversed(group)


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
