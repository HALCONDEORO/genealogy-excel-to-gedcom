# Gramps Web privacy remediation (after a GEDCOM import)

The code and its messages are in Spanish. This README is in English.

## Purpose

When you import a GEDCOM, `1 RESN privacy` lines are not turned into Gramps Web's `private` flag. They arrive as an inert attribute (type RESN, empty value) and every object stays `private=false`. These two tools derive the set of objects that should be private and mark them, with fail-closed checks.

- `priv_derive.py` derives and validates the target set from two independent routes, the source GED and the live API graph, which must agree object by object. It covers people with a RESN attribute, their events, families with at least one RESN member, notes at every level (person, event, family, family event, citation and source), citations (matched by the `| C... |` label in the page field) and the cited sources. Any disagreement aborts with exit code 1 and no targets. Any event content not covered by the identity key aborts instead of being ignored. Any list that reaches 2000 objects aborts (possible truncation). `priv_derive_v18_emisor.py` is a later emitter variant with its own test.
- `priv_remediate.py` applies the targets in two phases. Phase 1 validates every target before any write (exactly one match per gramps_id with matching handle, `content_sha256` of the derived object, content fingerprint for citations and sources). Phase 2 re-validates each object right before its PUT, changes only `private`, and reads it back, comparing every other field structurally. A per-object JSONL log keeps the pre-image sha256. The first anomaly stops the run with exit code 1.

## How the right object is matched

- People: by the gramps_id in the imported RESN attribute (GED set equals API set, exactly).
- Families: by content (multiset of member gramps_ids), never by family ID, since the import can renumber them.
- Events: selected by owner (event_ref_list of selected people and families), identified by a content key: owner, type, normalised date, place, exact note texts, number of citations, number of media. Any other event content (description, attributes, tags) must be empty or the run aborts.
- Notes: by owner and exact text. Citations: by the embedded `| C... |` label. Sources: by exact title via the selected citations.

## Ambiguous twins across the privacy boundary

The API keeps no independent GED-to-handle provenance, so a swap between content-identical objects cannot be detected by matching. The rule:

- Identical objects that all get the same treatment: reported as a declared residual (`ambiguity.events_same_side`).
- Identical objects whose owners are partly private and partly public: every member is made private (`origin: conservative_ambiguity`) and the group is reported redacted. Notes follow the same rule by text.
- If a member added by that expansion has citations or media, the derivation aborts rather than leave dependents public.
- An event with no owner inside a group, or a GED/API disagreement on a boundary-crossing class: fail closed.

Honest limit: matching guarantees the correct treatment per content class, not the physical identity of an object.

## Usage with the fictional fixtures in this repo

You need a Gramps Web 26.x stack (the official docker-compose is enough). Credentials go in the environment variables `GRAMPS_USER` and `GRAMPS_PASS`, never in arguments.

1. Import `fixtures/adv.ged` through the API (`POST /api/importers/ged/file`, multipart field `file`).
2. Derive: `python3 priv_derive.py fixtures/adv.ged http://localhost:5000`. Expected: `status OK` with 11 targets.
3. Apply: pipe the targets JSON to `priv_remediate.py`. A second run should report everything as already done (idempotent).
4. Other fixtures: `adv2_cross.ged` (three identical births, two private owners and one public) and `adv4_fam.ged` (two identical marriages in different families).
5. Check as a Guest user: marked objects disappear from lists and direct GETs return a redacted object. Marking only a person does not propagate to their event, family, citation or media, which is why the derived set covers every type.

## Tests

- `tests/test_priv_derive_unit.py`: offline, standard library only. Run `python3 tests/test_priv_derive_unit.py`.
- `tests/test_emitter_v18.py`: offline emitter test.
- `tests/test_integration_fixture.py`: runs against a disposable Gramps Web stack. It is skipped unless `GRAMPS_TEST_BASE` and test credentials are set. Never point it at real data.

## Known limits

- API only, it does not touch the database.
- Events with their own citations or media are not supported (abort, never partial privatisation).
- Shared places are not marked. Check separately if private notes or citations hang off places.
- In Gramps Web 26.10.1 a private media file returns 500 (not 403) to a low-privilege role. No bytes leak.
- The token endpoint is rate limited, so both tools retry with backoff.
- Swaps between identical objects on the same side of the boundary are undetectable and harmless for privacy.
