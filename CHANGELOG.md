# Changelog

The authoritative changelog is **GitHub Releases** —
<https://github.com/twobeass/VBAlidator/releases> — regenerated on
every `main` push by `python-semantic-release` from the Conventional
Commit history. This file is kept only for the pre-release phase-0–5
milestones; everything from `v1.0.0` onwards lives in the GitHub
Release pages.

The project follows [Conventional Commits] and Semver:
- `feat:` → minor bump
- `fix:` / `perf:` / `refactor:` → patch bump
- `chore:` / `ci:` / `docs:` / `test:` / `style:` → no release

[Keep a Changelog]: https://keepachangelog.com/en/1.1.0/
[Conventional Commits]: https://www.conventionalcommits.org/

## v1.0.0 onwards

See <https://github.com/twobeass/VBAlidator/releases>.

## Pre-release — Phase 5: CI/CD pipeline + full doc overhaul (foundation for v1.0.x)

### Added — Phase 5: maximum CI/CD pipeline + full doc overhaul
- `.github/workflows/release.yml` — python-semantic-release →
  GitHub Release → PyPI Trusted Publisher (OIDC). Nightly TestPyPI
  dev-release on every `main` push.
- `.github/workflows/docker.yml` — multi-arch (amd64, arm64) image
  build, Trivy scan (HIGH/CRITICAL fail), GHCR push + image smoke test.
- `.github/workflows/security.yml` — pip-audit, bandit, CodeQL.
- `.github/workflows/pr-quality.yml` — Conventional-Commit title +
  message lint, PR-size labelling.
- `.github/workflows/docs.yml` — MkDocs Material → GitHub Pages.
- `.github/workflows/roundtrip.yml` — nightly Windows + Office VBE
  cross-check; auto-files an issue on disagreement.
- `.github/workflows/nightly-coverage.yml` — daily Awesome-VBA score
  drift snapshot.
- `Dockerfile` (multi-stage, non-root, healthcheck) + `.dockerignore`.
- `mkdocs.yml` + `docs/index.md`, `docs/quickstart.md`,
  `docs/ai-integration.md`, `docs/ci-cd.md`, `docs/roadmap.md`,
  `docs/changelog.md`. Architecture / Usage / Configuration rewritten
  to reflect the Phase 0–4 reality.
- README full rewrite around `precheck()`, `--host`, `--roundtrip`,
  the 0–100 score, and the documented rule catalogue.
- `.github/CODEOWNERS`, `.github/FUNDING.yml`,
  `.github/pull_request_template.md`, structured issue forms
  (`feature.yml`, `false_positive.yml`).
- `.commitlintrc.json` for Conventional-Commit message linting.
- `pyproject.toml` extras: `[docs]` (mkdocs + mkdocs-material) and
  `[roundtrip]` (pywin32 on Windows). Semantic-release config block.
- `__version__` exposed at `src/__init__.__version__` and tracked by
  semantic-release alongside `pyproject.toml:project.version`.

## 0.1.x — Phase 0–4

See [docs/changelog.md](docs/changelog.md) for the milestone-by-milestone
breakdown.
