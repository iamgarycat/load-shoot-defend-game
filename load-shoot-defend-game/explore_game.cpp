#include <algorithm>
#include <array>
#include <cmath>
#include <iomanip>
#include <iostream>
#include <limits>
#include <map>
#include <string>
#include <tuple>
#include <vector>

using Vec = std::vector<double>;
using Mat = std::vector<Vec>;

struct Solution {
  double value = 0.0;
  Vec p;
  Vec q;
};

static bool solve_linear(std::vector<Vec> a, Vec b, Vec &x) {
  const int n = static_cast<int>(a.size());
  for (int i = 0; i < n; ++i) a[i].push_back(b[i]);
  for (int c = 0; c < n; ++c) {
    int piv = c;
    for (int r = c + 1; r < n; ++r)
      if (std::abs(a[r][c]) > std::abs(a[piv][c])) piv = r;
    if (std::abs(a[piv][c]) < 1e-11) return false;
    std::swap(a[piv], a[c]);
    double z = a[c][c];
    for (int j = c; j <= n; ++j) a[c][j] /= z;
    for (int r = 0; r < n; ++r) if (r != c) {
      double f = a[r][c];
      for (int j = c; j <= n; ++j) a[r][j] -= f * a[c][j];
    }
  }
  x.resize(n);
  for (int i = 0; i < n; ++i) x[i] = a[i][n];
  return true;
}

static std::vector<std::vector<int>> subsets(int n, int k) {
  std::vector<std::vector<int>> out;
  for (int mask = 0; mask < (1 << n); ++mask) {
    if (__builtin_popcount(static_cast<unsigned>(mask)) != k) continue;
    std::vector<int> s;
    for (int i = 0; i < n; ++i) if (mask & (1 << i)) s.push_back(i);
    out.push_back(s);
  }
  return out;
}

static Solution game_value(const Mat &m) {
  const int nr = static_cast<int>(m.size());
  const int nc = static_cast<int>(m[0].size());
  const double eps = 2e-8;
  for (int k = 1; k <= std::min(nr, nc); ++k) {
    for (const auto &is : subsets(nr, k)) for (const auto &js : subsets(nc, k)) {
      std::vector<Vec> ap(k + 1, Vec(k + 1));
      Vec bp(k + 1, 0.0), xp;
      for (int jj = 0; jj < k; ++jj) {
        for (int ii = 0; ii < k; ++ii) ap[jj][ii] = m[is[ii]][js[jj]];
        ap[jj][k] = -1.0;
      }
      for (int ii = 0; ii < k; ++ii) ap[k][ii] = 1.0;
      bp[k] = 1.0;
      if (!solve_linear(ap, bp, xp)) continue;

      std::vector<Vec> aq(k + 1, Vec(k + 1));
      Vec bq(k + 1, 0.0), xq;
      for (int ii = 0; ii < k; ++ii) {
        for (int jj = 0; jj < k; ++jj) aq[ii][jj] = m[is[ii]][js[jj]];
        aq[ii][k] = -1.0;
      }
      for (int jj = 0; jj < k; ++jj) aq[k][jj] = 1.0;
      bq[k] = 1.0;
      if (!solve_linear(aq, bq, xq)) continue;
      const double vp = xp[k], vq = xq[k];
      if (std::abs(vp - vq) > 1e-6) continue;
      bool ok = true;
      for (int i = 0; i < k; ++i)
        if (xp[i] < -eps || xq[i] < -eps) ok = false;
      Vec p(nr, 0.0), q(nc, 0.0);
      for (int i = 0; i < k; ++i) p[is[i]] = std::max(0.0, xp[i]);
      for (int j = 0; j < k; ++j) q[js[j]] = std::max(0.0, xq[j]);
      for (int j = 0; j < nc; ++j) {
        double z = 0;
        for (int i = 0; i < nr; ++i) z += p[i] * m[i][j];
        if (z < vp - eps) ok = false;
      }
      for (int i = 0; i < nr; ++i) {
        double z = 0;
        for (int j = 0; j < nc; ++j) z += m[i][j] * q[j];
        if (z > vp + eps) ok = false;
      }
      if (ok) return {vp, p, q};
    }
  }
  std::cerr << "no equilibrium\n";
  std::abort();
}

enum Action { L = 0, S = 1, D = 2 };

static std::vector<Action> actions(int r) {
  std::vector<Action> a{L};
  if (r >= 2) a.push_back(S);
  if (r % 2 == 1) a.push_back(D);
  return a;
}

static int next_r(int r, Action a) {
  if (a == L) return r + (r % 2 ? 2 : 3);
  if (a == S) return r - (r % 2 ? 2 : 1);
  return r - 1;
}

static Mat matrix_at(const std::vector<Vec> &prev, int r, int s) {
  auto ar = actions(r), ac = actions(s);
  Mat m(ar.size(), Vec(ac.size()));
  for (int i = 0; i < static_cast<int>(ar.size()); ++i) {
    for (int j = 0; j < static_cast<int>(ac.size()); ++j) {
      if (ar[i] == S && ac[j] == L) m[i][j] = 1.0;
      else if (ar[i] == L && ac[j] == S) m[i][j] = -1.0;
      else m[i][j] = prev[next_r(r, ar[i])][next_r(s, ac[j])];
    }
  }
  return m;
}

static std::string action_name(Action a) {
  if (a == L) return "L";
  if (a == S) return "S";
  return "D";
}

static bool forced_state(int r, int s) {
  return r / 2 >= s + 1 || s / 2 >= r + 1;
}

static double corrected_phi(int r, int s) {
  if (forced_state(r, s)) return 0.0;
  return std::max(10.0, static_cast<double>(r + s + (r % 2) + (s % 2)));
}

int main(int argc, char **argv) {
  int nmax = argc > 1 ? std::stoi(argv[1]) : 30;
  int target_max = 60;
  int print_max = 15;
  int rmax = target_max + 3 * nmax + 3;
  std::vector<Vec> lo(rmax + 1, Vec(rmax + 1, -1.0));
  std::vector<Vec> hi(rmax + 1, Vec(rmax + 1, 1.0));
  std::vector<Vec> ze(rmax + 1, Vec(rmax + 1, 0.0));
  double global_best_drift = -1e100;
  int global_n = -1, global_r = -1, global_s = -1;
  double global_best_ls = -1e100;
  int global_ls_n=-1, global_ls_r=-1, global_ls_s=-1;
  double global_bad_crossdiff=-1e100; int gcd_n=-1,gcd_r=-1,gcd_s=-1;
  double global_even_load=-1; int gel_n=-1,gel_r=-1,gel_s=-1;
  double global_k_violation=-1e100; int gk_n=-1,gk_r=-1,gk_s=-1;
  double global_adj_violation=-1e100; int ga_n=-1,ga_r=-1,ga_s=-1;
  double global_dom_violation=-1e100; int gd_n=-1,gd_r=-1,gd_s=-1;
  double global_dom2_violation=-1e100; int gd2_n=-1,gd2_r=-1,gd2_s=-1;
  for (int n = 1; n <= nmax; ++n) {
    int lim = rmax - 3 * n;
    std::vector<Vec> nlo(rmax + 1, Vec(rmax + 1));
    std::vector<Vec> nhi(rmax + 1, Vec(rmax + 1));
    std::vector<Vec> nze(rmax + 1, Vec(rmax + 1));
    for (int r = 0; r <= lim; ++r) for (int s = 0; s <= lim; ++s) {
      nlo[r][s] = game_value(matrix_at(lo, r, s)).value;
      nhi[r][s] = game_value(matrix_at(hi, r, s)).value;
      nze[r][s] = game_value(matrix_at(ze, r, s)).value;
    }
    lo.swap(nlo); hi.swap(nhi); ze.swap(nze);
    // Policies computed from these arrays are the first-step policies for
    // horizon n+1. Check the reserve drift on the fixed target box.
    for (int r=0;r<=target_max;++r) for(int s=0;s<=target_max;++s) {
      if(r+s<5 || r/2 >= s+1 || s/2 >= r+1) continue;
      auto ar=actions(r), ac=actions(s);
      auto sl=game_value(matrix_at(lo,r,s));
      auto sh=game_value(matrix_at(hi,r,s));
      double drift=0;
      for(int i=0;i<(int)ar.size();++i) for(int j=0;j<(int)ac.size();++j){
        double z=sh.p[i]*sl.q[j];
        if ((ar[i]==S&&ac[j]==L)||(ar[i]==L&&ac[j]==S)) drift -= z*(r+s);
        else drift += z*(next_r(r,ar[i])+next_r(s,ac[j])-(r+s));
      }
      if(drift>global_best_drift){global_best_drift=drift;global_n=n+1;global_r=r;global_s=s;}
      double pl=0, ps=0, ql=0, qs=0;
      for(int i=0;i<(int)ar.size();++i){if(ar[i]==L)pl=sh.p[i]; if(ar[i]==S)ps=sh.p[i];}
      for(int j=0;j<(int)ac.size();++j){if(ac[j]==L)ql=sl.q[j]; if(ac[j]==S)qs=sl.q[j];}
      double ls=(pl+ql)-(ps+qs);
      if(ls>global_best_ls){global_best_ls=ls;global_ls_n=n+1;global_ls_r=r;global_ls_s=s;}
      if(r>=2 && r%2==0 && s%2==1 && s>=3){
        double alpha=hi[r-1][s-2]-hi[r-1][s-1];
        double betap=hi[s-1][r-1]-hi[s-1][r+3];
        if(alpha-betap>global_bad_crossdiff){global_bad_crossdiff=alpha-betap;gcd_n=n;gcd_r=r;gcd_s=s;}
        double xx=0; for(int i=0;i<(int)ar.size();++i) if(ar[i]==L)xx=sh.p[i];
        if(xx>global_even_load){global_even_load=xx;gel_n=n+1;gel_r=r;gel_s=s;}
        double kval=2.0*(hi[r-1][s-2]-hi[r-1][s-1])-(1.0+hi[r+3][s-1]);
        if(kval>global_k_violation){global_k_violation=kval;gk_n=n;gk_r=r;gk_s=s;}
        double adj=(hi[r-1][s-2]-1.0)/2.0-hi[r-1][s-1];
        if(adj>global_adj_violation){global_adj_violation=adj;ga_n=n;ga_r=r;ga_s=s;}
        double dom=hi[r-1][s-2]-hi[r+3][s-1];
        if(dom>global_dom_violation){global_dom_violation=dom;gd_n=n;gd_r=r;gd_s=s;}
      }
      if(r+2<=lim && s+1<=lim){
        double dom2=hi[r][s]-hi[r+2][s+1];
        if(dom2>global_dom2_violation){global_dom2_violation=dom2;gd2_n=n;gd2_r=r;gd2_s=s;}
      }
    }
    if (n <= 20 || n % 5 == 0) {
      std::cerr << "seq " << n << ' ' << std::setprecision(12)
                << ze[5][2] << ' ' << ze[2][1] << ' ' << ze[5][3] << '\n';
      std::cerr << "bounds " << n << ' ' << std::setprecision(12)
                << lo[5][2] << ' ' << hi[5][2] << ' '
                << hi[5][2] - lo[5][2] << ' '
                << lo[2][1] << ' ' << hi[2][1] << ' '
                << hi[2][1] - lo[2][1] << '\n';
    }
    if (n == nmax) {
      std::cout << std::setprecision(12);
      double final_crossdiff=-1e100; int fc_r=-1,fc_s=-1;
      for(int rr=2;rr<=target_max;rr+=2) for(int ss=3;ss<=target_max;ss+=2){
        if(rr+ss<5 || rr/2>=ss+1 || ss/2>=rr+1) continue;
        double alpha=hi[rr-1][ss-2]-hi[rr-1][ss-1];
        double betap=hi[ss-1][rr-1]-hi[ss-1][rr+3];
        if(alpha-betap>final_crossdiff){final_crossdiff=alpha-betap;fc_r=rr;fc_s=ss;}
      }
      std::cout<<"final alpha-betap max="<<final_crossdiff<<" at="<<fc_r<<','<<fc_s<<'\n';
      double best_phi_drift=-1e100; int bpd_r=-1,bpd_s=-1;
      for(int rr=0;rr<=target_max;++rr) for(int ss=0;ss<=target_max;++ss){
        if(forced_state(rr,ss)) continue;
        auto aar=actions(rr), aac=actions(ss);
        auto slo=game_value(matrix_at(lo,rr,ss));
        auto shi=game_value(matrix_at(hi,rr,ss));
        double ed=-corrected_phi(rr,ss);
        for(int i=0;i<(int)aar.size();++i) for(int j=0;j<(int)aac.size();++j){
          double z=shi.p[i]*slo.q[j];
          if ((aar[i]==S&&aac[j]==L)||(aar[i]==L&&aac[j]==S)) continue;
          ed += z*corrected_phi(next_r(rr,aar[i]),next_r(ss,aac[j]));
        }
        if(ed>best_phi_drift){best_phi_drift=ed;bpd_r=rr;bpd_s=ss;}
      }
      std::cout<<"best corrected phi drift="<<best_phi_drift<<" at="<<bpd_r<<','<<bpd_s<<'\n';
      for(int rr=0;rr<=12;++rr) for(int ss=0;ss<=12;++ss){
        if(forced_state(rr,ss)) continue;
        double ph=corrected_phi(rr,ss);
        if(ph>12) continue;
        auto aar=actions(rr), aac=actions(ss);
        auto slo=game_value(matrix_at(lo,rr,ss));
        auto shi=game_value(matrix_at(hi,rr,ss));
        double ed=-ph;
        for(int i=0;i<(int)aar.size();++i) for(int j=0;j<(int)aac.size();++j){
          double z=shi.p[i]*slo.q[j];
          if ((aar[i]==S&&aac[j]==L)||(aar[i]==L&&aac[j]==S)) continue;
          ed += z*corrected_phi(next_r(rr,aar[i]),next_r(ss,aac[j]));
        }
        if(ed>-0.5) std::cout<<"phi detail "<<rr<<','<<ss<<" phi="<<ph<<" drift="<<ed<<'\n';
      }
      std::cout << "n=" << n << "\n";
      for (int r = 0; r <= print_max; ++r) {
        for (int s = 0; s <= print_max; ++s) {
          double gap = hi[r][s] - lo[r][s];
          if (gap > 1e-5) std::cout << "g(" << r << ',' << s << ")=" << gap << ' ';
        }
        std::cout << '\n';
      }
      for (auto [r,s] : std::vector<std::pair<int,int>>{{3,3},{5,2},{2,1},{5,3},{7,5},{5,4},{3,1}}) {
        auto sl = game_value(matrix_at(lo, r, s));
        auto sh = game_value(matrix_at(hi, r, s));
        std::cout << "next state " << r << ',' << s << " gap=" << sh.value-sl.value << "\n";
        auto ar=actions(r), ac=actions(s);
        std::cout << "  p_hi";
        for(int i=0;i<(int)ar.size();++i) std::cout << ' ' << action_name(ar[i]) << ':' << sh.p[i];
        std::cout << " q_lo";
        for(int j=0;j<(int)ac.size();++j) std::cout << ' ' << action_name(ac[j]) << ':' << sl.q[j];
        std::cout << '\n';
      }
      std::cout << "diagonal cross terminal probabilities for horizon n+1\n";
      for (int r = 0; r <= print_max; ++r) {
        auto sl = game_value(matrix_at(lo, r, r));
        auto sh = game_value(matrix_at(hi, r, r));
        auto a = actions(r);
        double term = 0.0;
        for (int i = 0; i < (int)a.size(); ++i)
          for (int j = 0; j < (int)a.size(); ++j)
            if ((a[i] == S && a[j] == L) || (a[i] == L && a[j] == S))
              term += sh.p[i] * sl.q[j];
        std::cout << "r=" << r << " term=" << term << " p";
        for (int i = 0; i < (int)a.size(); ++i)
          std::cout << ' ' << action_name(a[i]) << ':' << sh.p[i];
        std::cout << '\n';
      }
      std::cout << "zero-horizon odd diagonal parameters at n\n";
      for (int r = 3; r <= print_max; r += 2) {
        double aa = ze[r + 2][r - 1];
        double bb = ze[r - 1][r - 2];
        double z = 1 + aa + bb;
        std::cout << "r=" << r << " A=" << aa << " B=" << bb
                  << " pL=" << bb/z << " pS=" << aa/z << " pD=" << 1/z << '\n';
      }
      std::map<std::pair<int,int>, double> dist;
      dist[{5,2}] = 1.0;
      std::cout << "stationary cross survival from 5,2\n";
      for (int t = 1; t <= 20; ++t) {
        std::map<std::pair<int,int>, double> nd;
        double absorbed = 0.0;
        for (auto const &[x, mass] : dist) {
          auto [r,s] = x;
          auto ar=actions(r), ac=actions(s);
          auto sl=game_value(matrix_at(lo,r,s));
          auto sh=game_value(matrix_at(hi,r,s));
          for(int i=0;i<(int)ar.size();++i) for(int j=0;j<(int)ac.size();++j) {
            double z=mass*sh.p[i]*sl.q[j];
            if ((ar[i]==S && ac[j]==L)||(ar[i]==L&&ac[j]==S)) absorbed+=z;
            else nd[{next_r(r,ar[i]),next_r(s,ac[j])}]+=z;
          }
        }
        dist.swap(nd);
        double surv=0, mean=0;
        for(auto const &[x,mass]:dist){surv+=mass;mean+=mass*(x.first+x.second);}
        std::cout << "t="<<t<<" surv="<<surv<<" states="<<dist.size()
                  <<" conditional_sum="<<(surv?mean/surv:0)<<'\n';
      }
      std::cout << "cross reserve drift maxima by minimum total reserve\n";
      for (int cutoff : {0,5,10,20,30,40,60,80}) {
        double best = -1e9; int br=-1, bs=-1; double bt=0;
        for (int r=0;r<=target_max;++r) for(int s=0;s<=target_max;++s) {
          if(r+s<cutoff) continue;
          if(r/2 >= s+1 || s/2 >= r+1) continue;
          auto ar=actions(r), ac=actions(s);
          auto sl=game_value(matrix_at(lo,r,s));
          auto sh=game_value(matrix_at(hi,r,s));
          double drift=0, term=0;
          for(int i=0;i<(int)ar.size();++i) for(int j=0;j<(int)ac.size();++j){
            double z=sh.p[i]*sl.q[j];
            if ((ar[i]==S&&ac[j]==L)||(ar[i]==L&&ac[j]==S)) term+=z;
            else drift += z*(next_r(r,ar[i])+next_r(s,ac[j])-(r+s));
          }
          // Absorption is assigned zero future reserve, so include its Lyapunov drop.
          drift -= term*(r+s);
          if(drift>best){best=drift;br=r;bs=s;bt=term;}
        }
        std::cout<<"cut="<<cutoff<<" best="<<best<<" at="<<br<<','<<bs<<" term="<<bt<<'\n';
        if (cutoff==5 && br>=0) {
          auto ar=actions(br), ac=actions(bs);
          auto sl=game_value(matrix_at(lo,br,bs));
          auto sh=game_value(matrix_at(hi,br,bs));
          std::cout<<"  p";
          for(int i=0;i<(int)ar.size();++i)std::cout<<' '<<action_name(ar[i])<<':'<<sh.p[i];
          std::cout<<" q";
          for(int j=0;j<(int)ac.size();++j)std::cout<<' '<<action_name(ac[j])<<':'<<sl.q[j];
          std::cout<<'\n';
        }
      }
      std::cout << "detailed cross policies/drift in nonforced states <=12\n";
      for (int r=0;r<=12;++r) for(int s=0;s<=12;++s) {
        if(r+s<3 || r/2 >= s+1 || s/2 >= r+1) continue;
        auto ar=actions(r), ac=actions(s);
        auto sl=game_value(matrix_at(lo,r,s));
        auto sh=game_value(matrix_at(hi,r,s));
        double drift=0, term=0;
        for(int i=0;i<(int)ar.size();++i) for(int j=0;j<(int)ac.size();++j){
          double z=sh.p[i]*sl.q[j];
          if ((ar[i]==S&&ac[j]==L)||(ar[i]==L&&ac[j]==S)) {
            term += z;
            drift -= z*(r+s);
          } else {
            drift += z*(next_r(r,ar[i])+next_r(s,ac[j])-(r+s));
          }
        }
        std::cout<<"detail "<<r<<','<<s<<" val="<<sl.value<<','<<sh.value
                 <<" drift="<<drift<<" term="<<term<<" p";
        for(int i=0;i<(int)ar.size();++i)std::cout<<' '<<action_name(ar[i])<<':'<<sh.p[i];
        std::cout<<" q";
        for(int j=0;j<(int)ac.size();++j)std::cout<<' '<<action_name(ac[j])<<':'<<sl.q[j];
        std::cout<<'\n';
      }
    }
  }
  std::cerr << "global drift check horizon<= " << nmax+1
            << " box<= " << target_max << ": best=" << std::setprecision(12)
            << global_best_drift << " at horizon=" << global_n
            << " state=" << global_r << ',' << global_s << '\n';
  std::cerr << "global (loads-shots) max="<<global_best_ls<<" at horizon="<<global_ls_n
            <<" state="<<global_ls_r<<','<<global_ls_s<<'\n';
  std::cerr << "global alpha-betap max="<<global_bad_crossdiff<<" at array horizon="<<gcd_n
            <<" state="<<gcd_r<<','<<gcd_s<<'\n';
  std::cerr << "global even-vs-odd load max="<<global_even_load<<" at horizon="<<gel_n
            <<" state="<<gel_r<<','<<gel_s<<'\n';
  std::cerr << "global K violation="<<global_k_violation<<" at array horizon="<<gk_n
            <<" state="<<gk_r<<','<<gk_s<<'\n';
  std::cerr << "global adj violation="<<global_adj_violation<<" at array horizon="<<ga_n
            <<" state="<<ga_r<<','<<ga_s<<'\n';
  std::cerr << "global dominance violation="<<global_dom_violation<<" at array horizon="<<gd_n
            <<" state="<<gd_r<<','<<gd_s<<'\n';
  std::cerr << "global (+2,+1) dominance violation="<<global_dom2_violation<<" at array horizon="<<gd2_n
            <<" state="<<gd2_r<<','<<gd2_s<<'\n';
}
