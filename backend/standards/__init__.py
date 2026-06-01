"""Standards compliance rule-packs + checker for DOC-AI.

Turns "a better grammar checker" into "does this document qualify for safety
certification?" Each standard is a declarative JSON rule-pack (required sections,
naming conventions, FMEA table completeness) loaded at runtime, so new standards
can be added without code changes.
"""
