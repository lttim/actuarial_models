"""2017 CSO Ultimate mortality table loader (sex × smoker class).

Lightweight wrapper that reads a sex- and smoker-class-specific
``q_x`` CSV from :data:`data_registry.REGISTRY` and exposes the same
``monthly_survival_to_payment`` surface as
:class:`pricing_projection.MortalityTableQx`. The four life products
(WL / UL / IUL / VUL) default to this table; the existing annuity
products (SPIA / RILA / VA / MYGA / FIA) keep using their current
default mortality.

Synthetic placeholder warning
-----------------------------
The data files shipped under ``data/mortality/cso_2017_ult/`` are
**Gompertz-Makeham approximations** (see
``scripts/generate_cso_2017_synthetic.py``), NOT licensed CSO 2017
Ultimate data. Production users overlay their own licensed file at the
same path. The loader does not differentiate between the synthetic and
the licensed version -- both are valid CSV inputs -- so the swap is a
file-overlay, not a code change.

Reference: Section 1.4 of ``docs/seven_product_rollout_plan.md``.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Literal

import numpy as np

import pricing_projection as sp
from data_registry import get_artifact

Sex = Literal["male", "female"]
SmokerClass = Literal["nonsmoker", "smoker"]


_ARTIFACT_BY_KEY: dict[tuple[str, str], str] = {
    ("male", "nonsmoker"): "cso_2017_ult_male_nonsmoker_qx",
    ("female", "nonsmoker"): "cso_2017_ult_female_nonsmoker_qx",
    ("male", "smoker"): "cso_2017_ult_male_smoker_qx",
    ("female", "smoker"): "cso_2017_ult_female_smoker_qx",
}


@dataclass(frozen=True, slots=True)
class MortalityTable2017CSO:
    """2017 CSO Ultimate annual ``q_x`` for a specific sex × smoker cohort.

    Internally backed by :class:`pricing_projection.MortalityTableQx`
    so all downstream surfaces (monthly survival, qx_at_int_age, etc.)
    work identically.

    Attributes
    ----------
    sex:
        Either ``"male"`` or ``"female"``.
    smoker_class:
        Either ``"nonsmoker"`` or ``"smoker"``.
    table:
        The underlying :class:`MortalityTableQx`.
    """

    sex: Sex
    smoker_class: SmokerClass
    table: sp.MortalityTableQx

    @staticmethod
    def load(*, sex: Sex, smoker_class: SmokerClass) -> MortalityTable2017CSO:
        """Load the 2017 CSO Ultimate table for the requested cohort."""
        if sex not in ("male", "female"):
            raise ValueError(f"sex must be 'male' or 'female'; got {sex!r}")
        if smoker_class not in ("nonsmoker", "smoker"):
            raise ValueError(f"smoker_class must be 'nonsmoker' or 'smoker'; got {smoker_class!r}")
        artifact_name = _ARTIFACT_BY_KEY[(sex, smoker_class)]
        artifact = get_artifact(artifact_name)
        path = artifact.path
        if not path.exists():
            raise FileNotFoundError(
                f"CSO 2017 placeholder data not found at {path}. Run "
                "`python scripts/generate_cso_2017_synthetic.py` to "
                "regenerate, or overlay your licensed CSO file at that path."
            )
        table = sp.MortalityTableQx.load_qx_csv(str(path))
        return MortalityTable2017CSO(sex=sex, smoker_class=smoker_class, table=table)

    @property
    def ages(self) -> np.ndarray:
        return self.table.ages

    @property
    def qx(self) -> np.ndarray:
        return self.table.qx

    def qx_at_int_age(self, age_int: int) -> float:
        return self.table.qx_at_int_age(age_int)

    def monthly_survival_to_payment(
        self,
        *,
        issue_age: int,
        n_months: int,
        valuation_year: int | None = None,
    ) -> np.ndarray:
        """Forward to the underlying :class:`MortalityTableQx`.

        ``valuation_year`` is accepted for signature compatibility with
        :class:`MortalityTableRP2014MP2016` but is ignored (CSO 2017
        Ultimate is a static table — no calendar-year improvement).
        """
        del valuation_year
        return self.table.monthly_survival_to_payment(
            issue_age=int(issue_age),
            n_months=int(n_months),
        )


def load_2017_cso_ultimate(*, sex: Sex, smoker_class: SmokerClass) -> MortalityTable2017CSO:
    """Convenience free-function alias for :meth:`MortalityTable2017CSO.load`."""
    return MortalityTable2017CSO.load(sex=sex, smoker_class=smoker_class)


def cso_2017_artifact_path(*, sex: Sex, smoker_class: SmokerClass) -> Path:
    """Return the registry path for the requested CSO cohort."""
    artifact_name = _ARTIFACT_BY_KEY[(sex, smoker_class)]
    return get_artifact(artifact_name).path


__all__ = [
    "MortalityTable2017CSO",
    "Sex",
    "SmokerClass",
    "cso_2017_artifact_path",
    "load_2017_cso_ultimate",
]
