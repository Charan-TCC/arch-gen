"""archgen — spec-driven, editable-PPTX cloud architecture diagram generator.

Pipeline: prompt -> agent -> spec.yaml -> lint -> graphviz layout -> renderer -> out.pptx

The agent produces *only* a validated declarative spec (meaning); the renderer
owns all aesthetics (sizing, spacing, colors, arrows, legends) so every diagram
comes out consistent regardless of what the agent wrote.
"""

__version__ = "0.1.0"
