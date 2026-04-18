# Reproducible build / smoke environment for the actuarial models.
#
# Pinned base image: Python 3.12 on Debian Bookworm (matches CI ubuntu-latest +
# Python matrix). Pinned dependencies via annuity_model/requirements.lock.
#
# Build:  docker build -t annuity-model .
# Smoke:  docker run --rm annuity-model
# Tests:  docker run --rm annuity-model pytest -q
# Shell:  docker run --rm -it annuity-model bash

FROM python:3.12-slim-bookworm@sha256:3d77c6a48fcde98dbef0d33f0b3e95e8fc5b0b5d9c7d6c7e9c2c0a2f4e6f8a1b

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

ENV PYTHONPATH=/app/annuity_model

WORKDIR /app/annuity_model

# Default command runs the deep smoke (build + validate every product workbook).
CMD ["python", "scripts/deep_smoke.py"]
