# Gallery — committed reference renders

**Look here first.** These PNGs are the current rendered output of every example
spec (plus the bespoke azure_modern_data_platform board). They exist so you don't
have to re-render anything just to see what a spec produces. Open the image;
don't run the pipeline back-to-back to "check."

Regenerate after any change to the renderer, layout, theme, icons, or an example:

```bash
make gallery        # -> ./scripts/build_gallery.sh
```

If a diff changes the renderer/layout/theme, rebuild and eyeball these images —
that is the visual QA (DESIGN_GUIDELINES §"required visual QA"). A single
`make gallery` covers all examples at once.

| Image | Source | Profile |
|---|---|---|
| ![serverless_api](serverless_api.png) | [serverless_api.yaml](../../examples/serverless_api.yaml) | flow (graphviz) |
| ![event_driven_pipeline](event_driven_pipeline.png) | [event_driven_pipeline.yaml](../../examples/event_driven_pipeline.yaml) | flow (graphviz) |
| ![rag_chatbot](rag_chatbot.png) | [rag_chatbot.yaml](../../examples/rag_chatbot.yaml) | flow (graphviz) |
| ![sample_spec](sample_spec.png) | [sample_spec.yaml](../../examples/sample_spec.yaml) | flow (graphviz) |
| ![layered_data_platform](layered_data_platform.png) | [layered_data_platform.yaml](../../examples/layered_data_platform.yaml) | flow (graphviz, `layered`) |
| ![heimdall_architecture](heimdall_architecture.png) | [heimdall_architecture.yaml](../../examples/heimdall_architecture.yaml) | flow (graphviz, `layered`) |
| ![video_analytics](video_analytics.png) | [video_analytics.yaml](../../examples/video_analytics.yaml) | flow (graphviz, `layered`) |
| ![load_testing_spec](load_testing_spec.png) | [load_testing_spec.yaml](../../examples/load_testing_spec.yaml) | flow (graphviz, `layered`) |
| ![azure_private_aks](azure_private_aks.png) | [azure_private_aks.yaml](../../examples/azure_private_aks.yaml) | network |
| ![aws_discovery_platform](aws_discovery_platform.png) | [aws_discovery_platform.yaml](../../examples/aws_discovery_platform.yaml) | network (architecture board) |
| ![load_testing](load_testing.png) | [load_testing.yaml](../../examples/load_testing.yaml) | network |
| ![multi_region_azure](multi_region_azure.png) | [multi_region_azure.yaml](../../examples/multi_region_azure.yaml) | flow (graphviz) |
| ![azure_modern_data_platform](azure_modern_data_platform.png) | [build_azure_data_platform.py](../../custom_builds/build_azure_data_platform.py) | **bespoke** python-pptx (6-tier layered board; not the archgen renderer) |

`load_testing_spec.yaml` is scheduled for removal once its unique group-level
`border_color`/`fill_color`/`style` demonstration is ported onto another example
(see the project cleanup plan) — its row above will drop at that point.
