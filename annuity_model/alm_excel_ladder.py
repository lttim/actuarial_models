"""
Excel-native Treasury ladder (ALM_Engine). Mirrors ``run_alm_projection`` for
``rebalance_policy == liquidity_only``. Requires Excel 365 (LET).

See ``write_alm_engine_sheet`` for supported reinvest / disinvest / borrowing modes.
"""

from __future__ import annotations

from dataclasses import dataclass

import numpy as np
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

import spia_projection as sp

ALM_ENGINE_SHEET = "ALM_Engine"
ENGINE_HDR_ROW = 41
ENGINE_DATA_FIRST_ROW = 42


def _L(c: int) -> str:
    return get_column_letter(c)


def _excel_df_let(*, t_cell: str, y_last_row: int) -> str:
    y_rng = f"YieldCurve!$A$4:$A${y_last_row}"
    z_rng = f"YieldCurve!$B$4:$B${y_last_row}"
    return (
        f"LET(yt,{t_cell},yrng,{y_rng},zrng,{z_rng},sp,Inputs!$B$9,"
        f"br,IF(yt<=INDEX(yrng,1),1,IF(yt>=INDEX(yrng,ROWS(yrng)),ROWS(yrng)-1,MATCH(yt,yrng,1))),"
        f"lo,INDEX(yrng,br),hi,INDEX(yrng,br+1),zlo,INDEX(zrng,br),zhi,INDEX(zrng,br+1),"
        f"w,IF(hi=lo,0,(yt-lo)/(hi-lo)),ldf_lo,-(zlo+sp)*lo,ldf_hi,-(zhi+sp)*hi,"
        f"ldf,IF(yt<=INDEX(yrng,1),-(INDEX(zrng,1)+sp)*yt,"
        f"IF(yt>=INDEX(yrng,ROWS(yrng)),-(INDEX(zrng,ROWS(zrng))+sp)*yt,ldf_lo+w*(ldf_hi-ldf_lo))),"
        f"EXP(ldf))"
    )


@dataclass(frozen=True)
class ALMEngineLayout:
    first_data_row: int
    last_data_row: int
    col_mv_cash: int
    col_mv_bond_start: int
    col_debt_eom: int
    n_bonds: int


def write_alm_engine_sheet(
    ws,
    *,
    n_months: int,
    y_last_row: int,
    asm: sp.ALMAssumptions,
    initial_aum: float,
    snap_bucket_names: tuple[str, ...],
) -> ALMEngineLayout:
    if asm.rebalance_policy != "liquidity_only":
        raise ValueError(
            "Excel ALM ladder requires rebalance_policy='liquidity_only'. "
            "Select Hold-to-maturity bias in ALM settings for this export."
        )

    buckets = asm.allocation.buckets
    w = np.asarray(asm.allocation.weights, dtype=float)
    nb = len(buckets) - 1
    if nb < 1:
        raise ValueError("Need at least one bond bucket besides cash.")
    if tuple(str(b.name) for b in buckets) != snap_bucket_names:
        raise ValueError("Allocation buckets must match ALMExcelSnapshot.bucket_names.")

    nom = np.array([float(buckets[k + 1].tenor_years) for k in range(nb)], dtype=float)
    wsum_b = float(np.sum(w[1:]))
    wnorm = [float(w[k + 1]) / wsum_b for k in range(nb)] if wsum_b > 1e-15 else [1.0 / nb] * nb

    ws["A1"] = "ALM ladder engine (first principles — YieldCurve + Projection cashflows)"
    ws["A1"].font = Font(bold=True, size=12)
    ws["A2"] = (
        "Month recursion matches Python ALM when rebalance policy is liquidity_only. "
        "Uses LET for discount factors. Initial ladder from weights × AUM below."
    )
    ws.merge_cells("A2:H2")

    ws["A5"], ws["B5"] = "Initial AUM ($)", float(initial_aum)
    ws["A6"], ws["B6"] = "Δt (years)", 1.0 / 12.0
    ws["A7"], ws["B7"] = "Borrow: 1=scenario-linked rate, 0=fixed annual", int(
        1 if asm.borrowing_rate_mode == "scenario_linked" else 0
    )
    ws["A8"], ws["B8"] = (
        "Tenor (y) if scenario / fixed annual borrow rate (dec.)",
        float(asm.borrowing_rate_tenor_years if asm.borrowing_rate_mode == "scenario_linked" else asm.borrowing_rate_annual),
    )
    ws["A9"], ws["B9"] = "Borrow spread (dec.) if scenario-linked", float(asm.borrowing_spread_annual)
    ws["A10"], ws["B10"] = "1 = borrow before selling", int(1 if asm.borrowing_policy == "borrow_before_selling" else 0)
    ws["A11"], ws["B11"] = "1 = reinvest maturities pro_rata", int(1 if asm.reinvest_rule == "pro_rata" else 0)
    ws["A12"], ws["B12"] = "1 = disinvest shortest tenor first; 0 = pro_rata MV", int(
        1 if asm.disinvest_rule == "shortest_first" else 0
    )

    df_borrow = _excel_df_let(t_cell="$B$8", y_last_row=y_last_row)
    ws["A13"] = "Borrow rate (for exp(r*dt))"
    if asm.borrowing_rate_mode == "scenario_linked":
        ws["B13"] = f"=IF($B$7=1,MAX(0,-LN(({df_borrow}))/$B$8+$B$9),$B$8)"
    else:
        ws["B13"] = "=$B$8"

    w_row0 = 16
    ws["A15"] = "Cash weight w0"
    ws["B15"] = float(w[0])
    for k in range(nb):
        ws.cell(row=w_row0 + k, column=1, value=f"w bond {k + 1}").font = Font(bold=True)
        ws.cell(row=w_row0 + k, column=2, value=float(w[k + 1]))
        ws.cell(row=w_row0 + k, column=4, value=f"Nominal tenor (y) {k + 1}").font = Font(bold=True)
        ws.cell(row=w_row0 + k, column=5, value=float(nom[k]))

    # ---- Fixed column layout (1-based) ----
    c = 1
    col: dict[str, int | list[int]] = {}

    def take1(name: str) -> int:
        nonlocal c
        col[name] = c
        c += 1
        return col[name]

    def taken(name: str, n: int) -> list[int]:
        nonlocal c
        cols = list(range(c, c + n))
        col[name] = cols
        c += n
        return cols

    take1("mon")
    take1("cf")
    take1("d_acc")
    t_pm = taken("t_pm", nb)
    mat = taken("mat", nb)
    f_pm = taken("f_pm", nb)
    take1("cash_m")
    take1("rep1")
    take1("cash_r1")
    take1("debt_r1")
    df_pm = taken("df_pm", nb)
    mv_pm = taken("mv_pm", nb)
    take1("aum_re")
    take1("xsr")
    defc = taken("defc", nb)
    take1("dsum")
    split = taken("split", nb)
    dmv = taken("dmv", nb)
    take1("cash_re")
    f_re = taken("f_re", nb)
    t_re = taken("t_re", nb)
    take1("cash_cf")
    take1("need_raw")
    take1("cash_bb")
    take1("debt_bb")
    take1("need_dis")
    # DF at post-reinvest tenors (one LET per bond per row). Inlining these in disinvest
    # formulas exceeds Excel's 8192-char limit when nb or n_dis is moderately large.
    dfd = taken("dfd", nb)

    n_dis = nb + 2  # pro_rata disinvest may need > nb peels
    dis_need: list[int] = []
    dis_cash: list[int] = []
    dis_face: list[list[int]] = []
    for di in range(n_dis):
        dis_need.append(take1(f"nd{di}"))
        dis_cash.append(take1(f"cd{di}"))
        dis_face.append(taken(f"fd{di}", nb))

    take1("cash_pd")
    take1("debt_pb")
    take1("need_b2")
    take1("cash_br2")
    take1("debt_br2")
    take1("rep2")
    take1("cash_af2")
    take1("debt_af2")
    de = take1("de_e")
    ce = take1("ca_e")
    fe = taken("fe", nb)
    te = taken("te", nb)
    mv0 = take1("mv0")
    mvb = taken("mvb", nb)
    last_col = c - 1

    def C(name: str) -> int | list[int]:
        return col[name]

    hdr = ENGINE_HDR_ROW
    for cc in range(1, last_col + 1):
        ws.cell(row=hdr, column=cc, value=_L(cc)).font = Font(bold=True)

    R0 = ENGINE_DATA_FIRST_ROW
    WBASE = w_row0

    for i in range(n_months):
        r = R0 + i
        rp = r - 1
        mon = C("mon")
        if isinstance(mon, list):
            raise RuntimeError("bad col map")
        ws.cell(row=r, column=int(mon), value=1 if i == 0 else f"={_L(mon)}{rp}+1")

        cf_i = int(C("cf"))
        ws.cell(row=r, column=cf_i, value=f"=INDEX(Projection!$S:$S,3+{_L(mon)}{r})")

        d_acc = int(C("d_acc"))
        de_c = int(C("de_e"))
        ca_c = int(C("ca_e"))
        fe = C("fe")
        te = C("te")
        if not isinstance(fe, list) or not isinstance(te, list):
            raise RuntimeError("bad fe/te")

        debt_p = f"IF(ROW()={R0},0,{_L(de_c)}{rp})"
        cash_p = f"IF(ROW()={R0},$B$15*$B$5,{_L(ca_c)}{rp})"
        face_p = [
            (
                f"$B${WBASE + k}*$B$5/({_excel_df_let(t_cell=f'$E${WBASE + k}', y_last_row=y_last_row)})"
                if i == 0
                else f"{_L(fe[k])}{rp}"
            )
            for k in range(nb)
        ]
        trem_p = [f"$E${WBASE + k}" if i == 0 else f"{_L(te[k])}{rp}" for k in range(nb)]

        ws.cell(row=r, column=d_acc, value=f"=({debt_p})*EXP($B$13*$B$6)")

        for k in range(nb):
            ws.cell(row=r, column=t_pm[k], value=f"=MAX(0,{trem_p[k]}-$B$6)")

        m_parts = []
        for k in range(nb):
            fk = face_p[k]
            Tk = f"{_L(t_pm[k])}{r}"
            m_parts.append(f"IF(AND({fk}>1E-9,{Tk}<=1E-9),{fk},0)")
        msum = "+".join(m_parts)

        for k in range(nb):
            fk = face_p[k]
            Tk = f"{_L(t_pm[k])}{r}"
            ws.cell(row=r, column=mat[k], value=f"=IF(AND({fk}>1E-9,{Tk}<=1E-9),{fk},0)")

        for k in range(nb):
            fk = face_p[k]
            Tk = f"{_L(t_pm[k])}{r}"
            ws.cell(row=r, column=f_pm[k], value=f"=IF(AND({fk}>1E-9,{Tk}<=1E-9),0,{fk})")

        cash_m = int(C("cash_m"))
        ws.cell(row=r, column=cash_m, value=f"=({cash_p})+({msum})")

        rep1 = int(C("rep1"))
        cash_r1 = int(C("cash_r1"))
        debt_r1 = int(C("debt_r1"))
        ws.cell(row=r, column=rep1, value=f"=MIN({_L(cash_m)}{r},{_L(d_acc)}{r})")
        ws.cell(row=r, column=cash_r1, value=f"={_L(cash_m)}{r}-{_L(rep1)}{r}")
        ws.cell(row=r, column=debt_r1, value=f"={_L(d_acc)}{r}-{_L(rep1)}{r}")

        for k in range(nb):
            ws.cell(
                row=r,
                column=df_pm[k],
                value=f"={_excel_df_let(t_cell=f'{_L(t_pm[k])}{r}', y_last_row=y_last_row)}",
            )
        for k in range(nb):
            ws.cell(row=r, column=mv_pm[k], value=f"={_L(f_pm[k])}{r}*{_L(df_pm[k])}{r}")

        aum_re = int(C("aum_re"))
        mv_sum = "+".join(f"{_L(mv_pm[k])}{r}" for k in range(nb))
        ws.cell(row=r, column=aum_re, value=f"={_L(cash_r1)}{r}+{mv_sum}")

        xsr = int(C("xsr"))
        ws.cell(row=r, column=xsr, value=f"={_L(cash_r1)}{r}-$B$15*{_L(aum_re)}{r}")

        for k in range(nb):
            ws.cell(
                row=r,
                column=defc[k],
                value=f"=MAX(0,$B${WBASE + k}*{_L(aum_re)}{r}-{_L(mv_pm[k])}{r})",
            )
        dsum_i = int(C("dsum"))
        ds = "+".join(f"{_L(defc[j])}{r}" for j in range(nb))
        ws.cell(row=r, column=dsum_i, value=f"={ds}")

        for k in range(nb):
            wnk = wnorm[k]
            ws.cell(
                row=r,
                column=split[k],
                value=f"=IF({_L(dsum_i)}{r}>1E-9,{_L(defc[k])}{r}/{_L(dsum_i)}{r},{wnk})",
            )

        for k in range(nb):
            t_use = f"IF({_L(t_pm[k])}{r}>1E-9,{_L(t_pm[k])}{r},$E${WBASE + k})"
            denom = _excel_df_let(t_cell=f"({t_use})", y_last_row=y_last_row)
            wnk = wnorm[k]
            ws.cell(
                row=r,
                column=dmv[k],
                value=(
                    f"=IF(OR($B$11=0,{_L(xsr)}{r}<=1E-6),0,"
                    f"({_L(xsr)}{r}*IF({_L(dsum_i)}{r}>1E-9,{_L(defc[k])}{r}/{_L(dsum_i)}{r},{wnk}))/"
                    f"MAX(({denom}),1E-15))"
                ),
            )

        cash_re = int(C("cash_re"))
        dmv_sum = "+".join(f"{_L(dmv[j])}{r}" for j in range(nb))
        ws.cell(row=r, column=cash_re, value=f"=IF($B$11=0,{_L(cash_r1)}{r},{_L(cash_r1)}{r}-({dmv_sum}))")

        for k in range(nb):
            t_use = f"IF({_L(t_pm[k])}{r}>1E-9,{_L(t_pm[k])}{r},$E${WBASE + k})"
            denom = _excel_df_let(t_cell=f"({t_use})", y_last_row=y_last_row)
            ws.cell(
                row=r,
                column=f_re[k],
                value=(
                    f"=IF($B$11=0,{_L(f_pm[k])}{r},"
                    f"{_L(f_pm[k])}{r}+IF({_L(dmv[k])}{r}<=0,0,{_L(dmv[k])}{r}/"
                    f"MAX(({denom}),1E-15)))"
                ),
            )
        for k in range(nb):
            ws.cell(
                row=r,
                column=t_re[k],
                value=(
                    f"=IF($B$11=0,{_L(t_pm[k])}{r},"
                    f"IF(AND({_L(f_pm[k])}{r}<=1E-9,{_L(t_pm[k])}{r}<=1E-9),$E${WBASE + k},{_L(t_pm[k])}{r}))"
                ),
            )

        dfd_cols = C("dfd")
        if not isinstance(dfd_cols, list):
            raise RuntimeError("dfd")
        for k in range(nb):
            ws.cell(
                row=r,
                column=dfd_cols[k],
                value=f"={_excel_df_let(t_cell=f'{_L(t_re[k])}{r}', y_last_row=y_last_row)}",
            )

        ccf = int(C("cash_cf"))
        ws.cell(row=r, column=ccf, value=f"={_L(cash_re)}{r}-{_L(cf_i)}{r}")

        need_raw = int(C("need_raw"))
        ws.cell(row=r, column=need_raw, value=f"=MAX(0,-{_L(ccf)}{r})")

        c_bb = int(C("cash_bb"))
        d_bb = int(C("debt_bb"))
        need_dis = int(C("need_dis"))
        ws.cell(
            row=r,
            column=c_bb,
            value=f"=IF(AND($B$10=1,{_L(need_raw)}{r}>0),{_L(ccf)}{r}+{_L(need_raw)}{r},{_L(ccf)}{r})",
        )
        ws.cell(
            row=r,
            column=d_bb,
            value=f"=IF(AND($B$10=1,{_L(need_raw)}{r}>0),{_L(debt_r1)}{r}+{_L(need_raw)}{r},{_L(debt_r1)}{r})",
        )
        ws.cell(row=r, column=need_dis, value=f"=IF($B$10=1,0,{_L(need_raw)}{r})")

        tref = [f"{_L(t_re[k])}{r}" for k in range(nb)]
        df_dis = [f"{_L(dfd_cols[k])}{r}" for k in range(nb)]

        for ir in range(n_dis):
            nc, cc, fk_list = dis_need[ir], dis_cash[ir], dis_face[ir]
            if ir == 0:
                n0 = f"{_L(need_dis)}{r}"
                c0 = f"{_L(c_bb)}{r}"
                f0 = [f"{_L(f_re[k])}{r}" for k in range(nb)]
            else:
                pnc, pcc = dis_need[ir - 1], dis_cash[ir - 1]
                pfk = dis_face[ir - 1]
                n0 = f"{_L(pnc)}{r}"
                c0 = f"{_L(pcc)}{r}"
                f0 = [f"{_L(pfk[k])}{r}" for k in range(nb)]

            trnk = [f"({tref[k]}+{(k + 1)}*1E-9)" for k in range(nb)]
            tmin = f"MIN({','.join(f'IF({f0[k]}>1E-9,{trnk[k]},999)' for k in range(nb))})"
            mv_tot = "+".join(f"MAX(0,{f0[k]})*({df_dis[k]})" for k in range(nb))

            sells: list[str] = []
            for k in range(nb):
                sf = (
                    f"IF(AND($B$12=1,{f0[k]}>1E-9,ABS({trnk[k]}-({tmin}))<1E-9),"
                    f"MIN({f0[k]},{n0}/MAX(({df_dis[k]}),1E-15)),0)"
                )
                pr = (
                    f"IF(AND($B$12=0,{mv_tot}>1E-9,{f0[k]}>1E-9),"
                    f"MIN({f0[k]},{n0}*(MAX(0,{f0[k]})*({df_dis[k]}))/MAX(({mv_tot}),1E-15)),0)"
                )
                sells.append(f"IF($B$12=1,{sf},{pr})")

            pay = "+".join(f"({sells[k]})*({df_dis[k]})" for k in range(nb))
            ws.cell(row=r, column=nc, value=f"=MAX(0,{n0}-({pay}))")
            ws.cell(row=r, column=cc, value=f"={c0}+({pay})")
            for k in range(nb):
                ws.cell(row=r, column=fk_list[k], value=f"=MAX(0,{f0[k]}-({sells[k]}))")

        last_fk = dis_face[-1]
        last_cc = dis_cash[-1]
        c_pd = int(C("cash_pd"))
        d_pb = int(C("debt_pb"))
        ws.cell(row=r, column=c_pd, value=f"={_L(last_cc)}{r}")
        ws.cell(row=r, column=d_pb, value=f"={_L(d_bb)}{r}")

        need_b2 = int(C("need_b2"))
        ws.cell(row=r, column=need_b2, value=f"=MAX(0,-{_L(c_pd)}{r})")

        c_br2 = int(C("cash_br2"))
        d_br2 = int(C("debt_br2"))
        ws.cell(
            row=r,
            column=c_br2,
            value=f"=IF(AND($B$10=0,{_L(need_b2)}{r}>0),{_L(c_pd)}{r}+{_L(need_b2)}{r},{_L(c_pd)}{r})",
        )
        ws.cell(
            row=r,
            column=d_br2,
            value=f"=IF(AND($B$10=0,{_L(need_b2)}{r}>0),{_L(d_pb)}{r}+{_L(need_b2)}{r},{_L(d_pb)}{r})",
        )

        rep2 = int(C("rep2"))
        cash_af2 = int(C("cash_af2"))
        debt_af2 = int(C("debt_af2"))
        ws.cell(row=r, column=rep2, value=f"=MIN({_L(c_br2)}{r},{_L(d_br2)}{r})")
        ws.cell(row=r, column=cash_af2, value=f"={_L(c_br2)}{r}-{_L(rep2)}{r}")
        ws.cell(row=r, column=debt_af2, value=f"={_L(d_br2)}{r}-{_L(rep2)}{r}")

        for k in range(nb):
            ws.cell(row=r, column=fe[k], value=f"={_L(last_fk[k])}{r}")
        ws.cell(row=r, column=de_c, value=f"={_L(debt_af2)}{r}")
        ws.cell(row=r, column=ca_c, value=f"={_L(cash_af2)}{r}")
        for k in range(nb):
            ws.cell(row=r, column=te[k], value=f"={tref[k]}")
        mv0 = int(C("mv0"))
        mvb = C("mvb")
        if not isinstance(mvb, list):
            raise RuntimeError("mvb")
        ws.cell(row=r, column=mv0, value=f"={_L(ca_c)}{r}")
        for k in range(nb):
            ws.cell(
                row=r,
                column=mvb[k],
                value=f"={_L(fe[k])}{r}*({_excel_df_let(t_cell=f'{_L(te[k])}{r}', y_last_row=y_last_row)})",
            )

    last_r = R0 + n_months - 1
    mvb_l = C("mvb")
    if not isinstance(mvb_l, list):
        raise RuntimeError("mvb")
    return ALMEngineLayout(
        first_data_row=R0,
        last_data_row=last_r,
        col_mv_cash=int(C("mv0")),
        col_mv_bond_start=int(mvb_l[0]),
        col_debt_eom=int(C("de_e")),
        n_bonds=nb,
    )
