# cortex-ai-70b

Single home for the **CORTEX-70B** documentation and filing package: an autonomous,
self-learning 70B Mixture-of-Experts agent that runs locally, learns read-only, acquires
skills from physical USB drives, keeps a private `/memories` vault, and holds a persistent
self-sovereign identity. It builds on
[moe-async-engine](https://github.com/chrisfbaileycb-arch/moe-async-engine-.git) and
[transftran-ormers](https://github.com/chrisfbaileycb-arch/transftran-ormers).

This repository consolidates four previous repositories into one (see
[Consolidation history](#consolidation-history)).

## Layout

```
.
├── generators/
│   ├── gen_docs.py        # Technical Blueprint + Patent-Style Concept Summary (.docx)
│   └── build_filing.py    # Full filing document: "Same Thoughts · Read Skills" (.docx)
├── inputs/
│   ├── agent_deployment.txt        # Live agent deployment URL (arcada.app)
│   └── design_arena_artifact.txt   # DesignArena workspace artifact id for the export
├── requirements.txt
└── README.md
```

The `inputs/` files are provenance metadata from the DesignArena export; the generators
do not read them at runtime.

## Generating the documents

```bash
python3 -m venv .venv && . .venv/bin/activate
pip install -r requirements.txt

# Technical Blueprint + Patent-Style Concept Summary
python generators/gen_docs.py

# Filing document (also prints a paragraph/table/heading verification summary)
python generators/build_filing.py
```

Both scripts write their `.docx` output into the current working directory. Run them from
the repository root and the outputs will land there (they are git-ignored):

| Script | Output |
|---|---|
| `gen_docs.py` | `Cortex-70B-Technical-Blueprint.docx` |
| `gen_docs.py` | `Cortex-70B-Patent-Style-Concept-Summary.docx` |
| `build_filing.py` | `Cortex70B_Filing_Same_Thoughts_ReadOnly_USB_Memory_Identity_2026-09-07.docx` |

## Consolidation history

The following repositories were merged into this one. Their full git histories were
brought in with `git subtree add`, so `git log` here still shows every original commit.

| Former repository | What it held | Where it lives now |
|---|---|---|
| `cortex-ai-70b` (this repo) | `build_filing.py`, `input.txt` | `generators/build_filing.py`, `inputs/agent_deployment.txt` |
| `70b-cortez-new` | `gen_docs.py`, `input1.txt`, `input2.txt` | `generators/gen_docs.py`, `inputs/design_arena_artifact.txt`, `inputs/agent_deployment.txt` |
| `70b-cortez-new-2` | Byte-identical duplicate of `70b-cortez-new` | Deduplicated; history retained |
| `70b-blue-patent` | Empty (README only) | History retained |

`input2.txt` and `input.txt` were the same file and are kept once as
`inputs/agent_deployment.txt`.
