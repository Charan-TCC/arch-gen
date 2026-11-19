# archive/

Files kept for reference, not used by any build, test or doc.

- `unused-specs/` — `architecture.yaml`, `new_architecture.yaml`: specs moved out of the
  old `custom_builds/` folder; no build reads them.
- `outputs-legacy/` — renders from the old `outputs/` folder (Diageo harmonisation,
  recreated architecture). Their source specs no longer exist, so they can't be
  regenerated. Gitignored; local only.

`bespoke/gcp_proposals/build_v1.py` is the older GCP proposals build; `build.py` there
is the current one.
