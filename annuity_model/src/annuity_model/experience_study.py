"""Observed-vs-expected sample study for the actuarial-depth demo."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class ExperienceStudyRow:
    product_family: str
    cohort: str
    expected_claims: float
    observed_claims: float
    expected_lapses: float
    observed_lapses: float

    @property
    def claims_oe(self) -> float:
        return self.observed_claims / self.expected_claims if self.expected_claims else 0.0

    @property
    def lapse_oe(self) -> float:
        return self.observed_lapses / self.expected_lapses if self.expected_lapses else 0.0

    @property
    def review_flag(self) -> str:
        if abs(self.claims_oe - 1.0) >= 0.10 or abs(self.lapse_oe - 1.0) >= 0.15:
            return "Recommend assumption review"
        return "Within monitoring band"


SAMPLE_EXPERIENCE_STUDY: tuple[ExperienceStudyRow, ...] = (
    ExperienceStudyRow("Life", "Ages 35-49 nonsmoker", 1_850_000, 1_940_000, 11_200, 10_850),
    ExperienceStudyRow("Life", "Ages 50-64 smoker", 2_350_000, 2_725_000, 7_800, 7_610),
    ExperienceStudyRow("Annuity", "Ages 60-69 income", 4_900_000, 4_650_000, 2_050, 2_410),
    ExperienceStudyRow("Indexed", "ITM policy years 4-7", 1_150_000, 1_090_000, 4_600, 5_620),
)


def sample_experience_rows() -> list[dict[str, object]]:
    return [
        {
            "product_family": row.product_family,
            "cohort": row.cohort,
            "expected_claims": row.expected_claims,
            "observed_claims": row.observed_claims,
            "claims_oe": row.claims_oe,
            "expected_lapses": row.expected_lapses,
            "observed_lapses": row.observed_lapses,
            "lapse_oe": row.lapse_oe,
            "review_flag": row.review_flag,
        }
        for row in SAMPLE_EXPERIENCE_STUDY
    ]


__all__ = ["ExperienceStudyRow", "sample_experience_rows"]
