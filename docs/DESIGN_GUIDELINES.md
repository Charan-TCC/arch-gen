# Archgen Presentation Design Guidelines

These rules apply to every architecture slide produced by Archgen, whether the
specification is written by Codex, Claude, Antigravity, another agent, or a human.
The agent owns architectural meaning. The deterministic renderer owns layout and
visual styling.

## 1. Non-negotiable principles

1. Produce a clean, editable PowerPoint architecture slide, not a flattened image.
2. Keep the design flat and professional. Do not use shadows, glows, bevels,
   gradients, 3D effects, decorative reflections, or Office theme effects.
3. Do not invent components, connections, protocols, regions, availability zones,
   security controls, or implementation details that are absent from the input.
4. Do not deliver a slide with overlaps, clipped content, unresolved placeholders,
   broken connectors, hanging arrows, distorted icons, or unexpected text wrapping.
5. Prefer clarity over density. If the architecture cannot be read comfortably on
   one slide, simplify the labels, reduce low-value detail, or ask the user whether
   the architecture should be split across slides.

## 2. Logo and icon integrity

### Required asset behavior

- Use only an icon that exists in the Archgen asset registry.
- Use the exact official provider or technology asset whenever it is available.
- Never hallucinate, redraw, trace, approximate, or generate a brand logo.
- Never use an emoji, Unicode glyph, AI-generated image, homemade shape, or text
  monogram as though it were an official logo.
- Preserve every logo's aspect ratio. Scale uniformly and never stretch, squash,
  crop, rotate, recolor, or otherwise alter a brand asset.
- Only an asset explicitly marked `recolorable: true` may be tinted by the renderer,
  and then only to its own `brand_hex` (or the provider accent when unknown). Cloud
  brand packs (AWS/Azure/GCP) are never `recolorable`.
- Icons render in full color on a transparent background — never on a white box.
  Technology logos (`tech.*`) use real color artwork (Gil Barbara logos / Devicon),
  falling back to a Simple Icons monochrome logo tinted to its brand color. Generic
  concept icons (`generic.*`) are multi-color Flat Color Icons rendered as-is.
- Do not silently substitute one product's logo for a different product.

### Missing asset behavior

When an exact icon is unavailable, use this order:

1. Tell the user which asset is missing and ask whether a generic placeholder is
   acceptable when that decision would affect the meaning of the diagram.
2. If generation must continue, use a neutral, clearly generic placeholder from the
   registry, such as `generic.database`, `generic.queue`, or `generic.client`.
3. Keep the real service name as the label and record the substitution in the final
   response or asset report.

A placeholder must never imitate the missing brand logo. It must look intentionally
generic so that nobody mistakes it for official artwork.

## 3. Icon and label composition

- Treat each node as one atomic visual unit: card or container, icon, and label.
- Keep the icon and its label visually grouped and close together. A label must
  never float far from its icon or appear to belong to a neighboring node.
- Center the icon and label consistently unless the selected layout explicitly uses
  a left-aligned icon-label treatment.
- Keep a consistent gap between icon and label across all nodes of the same type.
- Use one icon per node. Do not stack multiple logos inside one icon slot.
- Where PowerPoint grouping is supported, group the card, icon, and label so they
  move as one unit. Otherwise, preserve the same bounding box and alignment contract
  so the elements behave visually as one unit.
- Container titles must remain inside their container and must not collide with
  child nodes, borders, or connectors.

## 4. Text and text-box rules

### Copy length

- Use short, audience-facing labels rather than sentences or implementation notes.
- Node labels must be 28 characters or fewer.
- Prefer one line. Allow at most two short lines for a node label.
- Declare every abbreviation in the specification so the abbreviation legend is
  complete and current.
- Do not expose agent instructions, planning notes, prompts, QA notes, or internal
  reasoning on the slide.

### Text-box sizing

- Size text boxes to the actual words plus consistent horizontal padding; do not
  create unnecessarily long or visibly empty text boxes.
- Keep the text box inside its parent card or container at all times.
- A one-line title or banner must never wrap to a second line.
- Do not shrink fonts to force lengthy content into a small box. Shorten the wording
  or change the layout first.
- Do not let PowerPoint auto-grow a text box beyond its assigned card, container, or
  slide boundary.
- Use consistent internal padding. Text must not touch the box border or appear
  vertically off-center.
- Inspect rendered output for font substitution, clipping, orphaned words, awkward
  line breaks, and unexpected wrapping.

## 5. Typography and hierarchy

- Use the theme font consistently; the current Archgen default is Segoe UI.
- Maintain a clear hierarchy: slide title, container title, node label, edge label,
  and legend text must be visually distinct.
- Keep equivalent items at the same font size and weight.
- Use bold only for meaningful hierarchy, not for decoration.
- Maintain sufficient contrast between text and its background.
- If no supplied template overrides the project theme, use the sizes defined in
  `archgen/render/theme.yaml` and never reduce them merely to fit excessive copy.

## 6. Connectors and arrows

- Every connector must have a valid source and destination node.
- Every connector endpoint must touch the border or defined anchor point of its
  source or destination. No connector may stop in empty space.
- No hanging arrows are allowed. If a flow continues outside the shown system, add
  a labeled external or boundary node and connect the arrow to that node.
- Arrowheads must point toward the destination and remain visible.
- Draw connectors before nodes so lines sit behind cards, icons, and labels.
- Route lines around nodes and labels. Lines must not cut through icons, text, or
  unrelated containers.
- Use orthogonal (right-angle) routing. Flow diagrams use native PowerPoint elbow
  connectors; network diagrams use explicit Manhattan polylines with semantic lanes.
  Minimize bends, crossings, parallel-line ambiguity, and long detours.
- Use connection colors and line styles consistently with their semantic meaning.
- Do not use an unlabeled color or line treatment that is absent from the legend.
- Edge labels must be short, centered near the relevant segment, and placed on a
  clean backing so the line does not reduce readability.

## 7. Layout, spacing, and grouping

- Use a consistent grid and equal outer margins.
- Keep related components inside a clearly labeled group or boundary.
- Use nesting only when it communicates a real architectural relationship.
- Do not overlap group containers, nodes, icons, labels, legends, or connectors.
- Keep spacing consistent among peer nodes and peer containers.
- Leave enough internal padding for container titles and child nodes.
- Keep provider boundaries and network boundaries visually distinct without adding
  decorative chrome.
- For an ordered tier architecture (source → ingest → store → serve → …), declare
  `layered: true` on the spec — a **structural** statement, not a styling one. The
  renderer then packs the groups into parallel bands and derives card size from tier
  density (the spec still names no coordinates, sizes, or colors).
- Do not create dense dashboard-style card grids, pills, tabs, badges, or controls.
  This is an architecture diagram, not an application interface.
- Use whitespace to clarify relationships, but avoid large unexplained empty regions.
- Legends must not cover the architecture and must be derived from current content.

## 8. Color and visual treatment

- Use only theme and provider palette colors defined by the project.
- Keep the slide background white or near-white. Use very pale provider tints only
  to clarify a real group, tier, cloud, network, or trust boundary.
- Select foreground colors against the actual background on which they appear; do
  not validate a color against white if it will be placed on a tinted panel.
- Keep backgrounds neutral and low-saturation so official logos, labels, and
  connectors remain the visual focus.
- Do not use color as decoration. Every semantic color must have a consistent meaning.
- Never rely on color alone. Pair semantic colors with labels, line styles, arrow
  direction, numbering, grouping, or a legend.
- Preserve official logo colors unless the asset is explicitly safe to recolor.
- Do not use transparency or low contrast that makes boundaries, text, or connectors
  difficult to see when projected.
- Maintain a flat visual treatment with no shadows or inherited Office effects.

### Accessibility checks

- Target WCAG 2.2 AA contrast as the project baseline, even though the output is a
  presentation rather than a web page.
- Normal text must have at least **4.5:1** contrast against its background.
- Large text may use a minimum of **3:1** contrast against its background.
- Meaningful icons, connector lines, arrowheads, and graphical boundaries should
  have at least **3:1** contrast against adjacent colors when they are required to
  understand the diagram.
- Do not recolor an official logo merely to make it pass a contrast check. Instead,
  place it on a permitted neutral background or use an approved light/dark asset
  variant when one exists.
- Test the rendered slide in grayscale and with common color-vision-deficiency
  simulations. Meanings must remain understandable without hue recognition.
- Inspect the slide at normal presentation size and in projection-like conditions;
  thin lines and pale text that disappear at a distance are not acceptable.

### Borders and strokes

- Use thin, visually light strokes. Default node borders should be **0.75–1 pt**;
  group and outer-boundary strokes should normally be **1–1.25 pt**.
- Use muted neutral or provider-aligned border colors rather than heavy black boxes.
- Use a darker or thicker stroke only when it communicates hierarchy, an outer trust
  boundary, or a selected semantic state.
- If a border is the only cue that a meaningful region exists, it must remain
  distinguishable from both adjoining backgrounds. Otherwise, supplement a pale
  border with a subtle fill and a clear boundary title.
- Keep equivalent boundaries at the same stroke weight. Do not mix arbitrary border
  thicknesses or use thick borders to compensate for poor spacing.
- Connector strokes must remain readable and consistent, but should not visually
  overpower icons, labels, or container boundaries.

## 9. Legends, numbering, and abbreviations

Legends must be generated from the current specification, positioned outside the
main architecture flow, and omitted when they add no useful information.

### Color and connection legends

- Add a connection legend whenever two or more colors, dash styles, or line
  treatments encode different connection meanings.
- Each legend sample must exactly match the rendered connector color, width, dash
  style, and meaning.
- Add a provider-color legend when multiple provider colors are used or when color
  encodes ownership. A clearly labeled single-provider boundary normally does not
  need a redundant provider legend.
- Never introduce a semantic color that is missing from the legend.
- Do not add a connection-color legend when every edge is the same ordinary traffic
  type and the meaning is already obvious from labels and arrow direction.

### Numbered flows

- Use numbered arrows only when sequence or reading order matters.
- Number steps consecutively from `1` to `N`, with no duplicate or missing numbers.
- Place each number next to the relevant connector segment without covering the
  arrow, node, icon, or label.
- When numbered steps are not self-explanatory, add a compact numbered-flow legend
  that maps each number to a short audience-facing action.
- Keep the numbering legend in the same visual order as the flow. Do not make the
  reader jump randomly around the slide.
- Do not use numbering merely as decoration. Remove it when the architecture is
  non-sequential or the arrows already communicate the flow clearly.

### Abbreviation legends

- Add an abbreviation legend whenever an abbreviation appears and its expansion is
  needed by the intended audience.
- Every displayed abbreviation must be declared once and expanded accurately.
- Sort abbreviations consistently, preferably alphabetically, unless first-use order
  is materially easier to follow.
- Omit the abbreviation legend when the slide contains no abbreviations.
- If a legend becomes long, simplify the labels or move supporting terminology to a
  separate slide rather than shrinking the legend into unreadable text.
- Keep provider, connection, numbered-flow, and abbreviation legends visually
  distinct, but style them as one restrained legend system.

## 10. Content fidelity

- The agent converts the user's description into architectural meaning only.
- The agent must not write PowerPoint drawing code or specify arbitrary coordinates,
  colors, font sizes, or decorative styles in the architecture specification.
- Every node, group, edge, and abbreviation must be supported by the user's input or
  by an explicitly stated assumption.
- If the input is ambiguous in a way that materially changes the architecture, ask
  the user instead of guessing.
- Keep labels audience-facing. Put detailed explanations in speaker notes or a
  separate document only when the user requests them.

## 11. Required validation and visual QA

Before delivery:

1. Rebuild the icon manifest after asset changes.
2. Lint the architecture specification and fix every reported error.
3. Run the complete automated test suite.
4. Render the final PowerPoint.
5. Run slide-boundary and overflow checks.
6. Render every final slide to an image.
7. Inspect each slide individually at full size, not only as a montage.
8. Confirm all icons are correct and undistorted.
9. Confirm all text is readable, concise, correctly wrapped, and inside its box.
10. Confirm every arrow is attached at both ends and points to the correct destination.
11. Run contrast checks for each text/background pair and every meaningful graphical
    object; verify the slide remains understandable without color alone.
12. Confirm borders are thin, consistent, and light without becoming invisible.
13. Confirm color, connection, numbered-flow, provider, and abbreviation legends are
    present only when needed and exactly match the rendered slide.
14. Confirm there are no overlaps, clipped objects, broken routes, theme shadows, or
    unresolved placeholders.
15. Open the `.pptx` and confirm that cards, labels, arrows, legends, and containers
    remain editable PowerPoint objects.

## 12. Delivery disclosure

The final response must identify:

- the final PowerPoint file;
- any generic placeholder used for a missing icon;
- any assumption that materially affects the diagram;
- any validation or visual-QA step that could not be completed.

Do not describe a slide as complete when a required asset is missing, a test fails,
or the rendered slide has not been visually inspected.

## 13. Acceptance checklist

A slide is ready only when every answer below is **yes**:

- Are all logos sourced from registered assets?
- Are missing logos disclosed and represented only by neutral placeholders?
- Are all icons uniformly scaled and undistorted?
- Are icons and labels visually grouped as one node?
- Are labels concise and text boxes fitted to their content?
- Are fonts readable without last-minute shrinking?
- Does every text/background pair meet the required contrast target?
- Are meaningful graphical objects distinguishable without relying on color alone?
- Are borders thin, light, consistent, and still visible at presentation size?
- Are all connectors attached at both ends with visible arrowheads?
- Are connectors routed behind nodes and away from text?
- Are connection-color, provider, numbered-flow, and abbreviation legends included
  only when needed and synchronized with the slide?
- If flow numbering is used, is it consecutive, unambiguous, and explained when
  necessary?
- Are all groups, nodes, labels, connectors, and legends free of overlap?
- Is the slide flat, consistent, and free of shadows and decorative effects?
- Did linting, tests, rendering, overflow checks, and full-size visual inspection pass?
- Is the final PowerPoint editable?
