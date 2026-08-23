from fractions import Fraction as F
from itertools import combinations
from functools import lru_cache
import sys

L, S, D = 0, 1, 2

def actions(r):
    z = [L]
    if r >= 2:
        z.append(S)
    if r & 1:
        z.append(D)
    return z

def nxt(r, a):
    if a == L:
        return r + (2 if r & 1 else 3)
    if a == S:
        return r - (2 if r & 1 else 1)
    return r - 1

def solve_linear(A, b):
    n = len(A)
    z = [list(A[i]) + [b[i]] for i in range(n)]
    for c in range(n):
        p = next((i for i in range(c, n) if z[i][c]), None)
        if p is None:
            return None
        z[c], z[p] = z[p], z[c]
        q = z[c][c]
        z[c] = [x / q for x in z[c]]
        for i in range(n):
            if i == c:
                continue
            q = z[i][c]
            if q:
                z[i] = [z[i][j] - q * z[c][j] for j in range(n + 1)]
    return [z[i][-1] for i in range(n)]

def row_lp(M):
    """Exact row maximin LP. Return (value, one optimal row distribution)."""
    m, n = len(M), len(M[0])
    # Variables p[0:m], v.  Equality sum(p)=1.  A vertex has m further
    # active constraints chosen among n payoff constraints and m p_i=0.
    best = None
    constraints = [('pay', j) for j in range(n)] + [('zero', i) for i in range(m)]
    for active in combinations(constraints, m):
        A = [[F(1) for _ in range(m)] + [F(0)]]
        b = [F(1)]
        for kind, k in active:
            if kind == 'pay':
                A.append([M[i][k] for i in range(m)] + [F(-1)])
                b.append(F(0))
            else:
                row = [F(0) for _ in range(m + 1)]
                row[k] = F(1)
                A.append(row)
                b.append(F(0))
        x = solve_linear(A, b)
        if x is None:
            continue
        p, v = x[:m], x[m]
        if any(t < 0 for t in p):
            continue
        if any(sum(p[i] * M[i][j] for i in range(m)) < v for j in range(n)):
            continue
        if best is None or v > best[0]:
            best = (v, p)
    if best is None:
        raise RuntimeError('LP vertex enumeration failed')
    return best

def game_matrix(prev, r, s):
    ar, ac = actions(r), actions(s)
    M = []
    for a in ar:
        row = []
        for b in ac:
            if a == S and b == L:
                row.append(F(1))
            elif a == L and b == S:
                row.append(F(-1))
            else:
                row.append(prev(nxt(r, a), nxt(s, b)))
        M.append(row)
    return M

def build(cutoff):
    @lru_cache(None)
    def V(n, r, s):
        if n == 0:
            return F(cutoff)
        M = game_matrix(lambda x, y: V(n - 1, x, y), r, s)
        return row_lp(M)[0]
    return V

def fmt(x):
    return f'{float(x):.12f} ({x.numerator.bit_length()}/{x.denominator.bit_length()} bits)'

def policy_outer_vertices(Mhi, value_lower):
    """Vertices of {p>=0, sum p=1, p^T Mhi[:,j]>=value_lower}."""
    m, n = len(Mhi), len(Mhi[0])
    constraints = [('zero', i) for i in range(m)] + [('pay', j) for j in range(n)]
    out = []
    for active in combinations(constraints, m - 1):
        A = [[F(1) for _ in range(m)]]
        b = [F(1)]
        for kind, k in active:
            if kind == 'zero':
                row = [F(0) for _ in range(m)]
                row[k] = F(1)
                A.append(row); b.append(F(0))
            else:
                A.append([Mhi[i][k] for i in range(m)])
                b.append(value_lower)
        p = solve_linear(A, b)
        if p is None or any(x < 0 for x in p):
            continue
        if any(sum(p[i] * Mhi[i][j] for i in range(m)) < value_lower for j in range(n)):
            continue
        if p not in out:
            out.append(p)
    return out

def forced(r, s):
    return r // 2 >= s + 1 or s // 2 >= r + 1

def phi(r, s):
    if forced(r, s):
        return F(0)
    h = r + s + (r & 1) + (s & 1)
    return F(max(10, h))

def drift_matrix(r, s):
    ar, ac = actions(r), actions(s)
    M = []
    for a in ar:
        row = []
        for b in ac:
            if (a == S and b == L) or (a == L and b == S):
                row.append(-phi(r, s))
            else:
                row.append(phi(nxt(r, a), nxt(s, b)) - phi(r, s))
        M.append(row)
    return M

def escape_matrix(r, s):
    ar, ac = actions(r), actions(s)
    cur = phi(r, s)
    M = []
    for a in ar:
        row = []
        for b in ac:
            if (a == S and b == L) or (a == L and b == S):
                row.append(F(1))
            else:
                rr, ss = nxt(r, a), nxt(s, b)
                row.append(F(1) if forced(rr, ss) or phi(rr, ss) != cur else F(0))
        M.append(row)
    return M

def extrema_bilinear(A, pv, qv):
    vals = [(sum(p[i] * A[i][j] * q[j] for i in range(len(p)) for j in range(len(q))), p, q)
            for p in pv for q in qv]
    return min(vals, key=lambda z: z[0]), max(vals, key=lambda z: z[0])

def verify_core(N):
    hi, lo = build(1), build(-1)
    bad = []
    records = []
    escapes = []
    poly = {}
    for r in range(13):
        for s in range(13):
            if forced(r, s) or phi(r, s) > 12:
                continue
            Mx = game_matrix(lambda x, y: hi(N - 1, x, y), r, s)
            My = game_matrix(lambda x, y: hi(N - 1, x, y), s, r)
            px = policy_outer_vertices(Mx, lo(N, r, s))
            py = policy_outer_vertices(My, lo(N, s, r))
            if not px or not py:
                raise RuntimeError(('empty outer polytope', r, s))
            poly[(r, s)] = px
            H = drift_matrix(r, s)
            _, mx = extrema_bilinear(H, px, py)
            records.append((mx[0], r, s, len(px), len(py)))
            if phi(r, s) == 10:
                mn_escape, _ = extrema_bilinear(escape_matrix(r, s), px, py)
                escapes.append((mn_escape[0], r, s))
            if mx[0] > 0:
                bad.append((r, s, mx[0], mx[1], mx[2]))
    records.sort(reverse=True)
    print('core N', N, 'states', len(records), 'bad', len(bad))
    for z in records[:20]:
        print('maxdrift', float(z[0]), z[1:])
    if bad:
        print('BAD')
        for z in bad:
            print(z[0:3], [float(x) for x in z[3]], [float(x) for x in z[4]])
    escapes.sort()
    print('one-step escape lower bounds at phi=10')
    for z in escapes:
        print(float(z[0]), z[1], z[2])
    for level in (F(10), F(12)):
        states = [x for x in poly if phi(*x) == level]
        e = {x: F(0) for x in states}
        for steps in range(1, 7):
            ne = {}
            for r, s in states:
                ar, ac = actions(r), actions(s)
                A = []
                for a in ar:
                    row = []
                    for b in ac:
                        if (a == S and b == L) or (a == L and b == S):
                            row.append(F(1)); continue
                        y = (nxt(r, a), nxt(s, b))
                        if forced(*y) or phi(*y) != level:
                            row.append(F(1))
                        else:
                            row.append(e[y])
                    A.append(row)
                mn, _ = extrema_bilinear(A, poly[(r, s)], poly[(s, r)])
                ne[(r, s)] = mn[0]
            e = ne
            print('level', int(level), 'steps', steps, 'min_escape', float(min(e.values())))
            if level == 10 and steps == 4:
                assert min(e.values()) > F(9, 100)
                print('CERTIFIED level 10, 4 steps: min_escape > 9/100')
            if level == 12 and steps == 2:
                assert min(e.values()) > F(1, 2)
                print('CERTIFIED level 12, 2 steps: min_escape > 1/2')
    return not bad

if __name__ == '__main__':
    N = int(sys.argv[1]) if len(sys.argv) > 1 else 8
    if len(sys.argv) > 2 and sys.argv[2] == 'verify':
        verify_core(N)
        raise SystemExit
    hi, lo = build(1), build(-1)
    states = [(3,3),(3,5),(5,3),(3,7),(5,5),(7,3),(3,6),(4,5),(5,4),(6,3),
              (5,2),(2,1)]
    for x in states:
        print(N, x, fmt(lo(N,*x)), fmt(hi(N,*x)), 'gap', float(hi(N,*x)-lo(N,*x)))
