# syntax=docker/dockerfile:1.7

# ---------- builder ------------------------------------------------------
FROM python:3.12-slim AS builder

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1

WORKDIR /build

# Copy only what's needed to build the wheel.
COPY pyproject.toml README.md ./
COPY src ./src

RUN pip install --upgrade pip build \
 && python -m build --wheel --sdist --outdir /wheels


# ---------- runtime ------------------------------------------------------
FROM python:3.12-slim AS runtime

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    VBALIDATOR_VERSION=unknown

LABEL org.opencontainers.image.title="VBAlidator" \
      org.opencontainers.image.description="Premium VBA static analyser & compile-safety prechecker" \
      org.opencontainers.image.source="https://github.com/twobeass/VBAlidator" \
      org.opencontainers.image.licenses="MIT"

# Run as a non-root user.
RUN groupadd --system --gid 1001 vba \
 && useradd --system --uid 1001 --gid vba --create-home --home /home/vba vba

WORKDIR /workspace

COPY --from=builder /wheels /wheels
RUN pip install --no-cache-dir --upgrade pip \
 && pip install --no-cache-dir /wheels/*.whl \
 && rm -rf /wheels /root/.cache

USER vba

# Mount the project to scan as `/workspace`.
VOLUME ["/workspace"]

# Lightweight self-test so health-checks can verify the binary works.
HEALTHCHECK --interval=1m --timeout=10s --retries=3 \
  CMD vbalidator --help > /dev/null || exit 1

ENTRYPOINT ["vbalidator"]
CMD ["--help"]
