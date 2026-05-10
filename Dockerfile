# Reproducible build / smoke environment for the actuarial models.
#
# Pinned base image: Python 3.12 on Debian Bookworm (matches CI ubuntu-latest +
# Python matrix). Pinned dependencies via annuity_model/requirements.lock.
#
# Build:  docker build -t annuity-model .
# Smoke:  docker run --rm annuity-model
# Tests:  docker run --rm annuity-model pytest -q
# Shell:  docker run --rm -it annuity-model bash

FROM python:3.12-slim-bookworm@sha256:d97792894a6a4162cae14da44542a83c75e56c77a27b92d58f3f83b7bc961292

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1

WORKDIR /app

# Install runtime + dev dependencies first so layer cache is reusable.
COPY annuity_model/requirements.lock annuity_model/requirements-dev.txt ./
RUN pip install --upgrade pip \
 && pip install -r requirements.lock -r requirements-dev.txt

# Copy the rest of the codebase.
COPY annuity_model/ ./annuity_model/
COPY actuarial_parity_kit/ ./actuarial_parity_kit/
COPY README.md AGENTS.md CONTRIBUTING.md ./

RUN pip install --no-build-isolation -e ./annuity_model

ENV PYTHONPATH=/app/annuity_model/src

WORKDIR /app/annuity_model

# Default command runs the deep smoke (build + validate every product workbook).
CMD ["python", "scripts/deep_smoke.py"]
