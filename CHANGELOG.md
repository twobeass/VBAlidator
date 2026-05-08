# Changelog

All notable changes to this project are documented here. The format is
loosely based on [Keep a Changelog] and the project follows
[Conventional Commits] / Semver via `python-semantic-release`.

[Keep a Changelog]: https://keepachangelog.com/en/1.1.0/
[Conventional Commits]: https://www.conventionalcommits.org/

Once the first semantic-release tag lands this file is regenerated
automatically from commit history. Until then it tracks the major
milestones manually — see also [docs/changelog.md](docs/changelog.md).

## [Unreleased]

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
