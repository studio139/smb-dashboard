"""SMB architecture-studio financial dashboard engine.

Deterministic pipeline: loader -> validator -> metrics -> (targets) -> generator/preview.
The code computes every number; the generator builds every layout. Nothing is invented.
"""
