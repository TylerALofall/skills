# Runtime Flow Questions

## Where the response and action come from

In the current scaffold, the model response is a dry-run nub in `assets/c-shell-template/main.c`:

1. `assemble_prompt_packet` builds the screen packet from the user prompt, screenshot path, mouse position, and `BTN_XX` button placeholders.
2. `local_model_stream_dry_run` is the single replacement point for the real local model runtime. Today it prints the packet and returns a sample model line: `ACTION hover BTN_01`.
3. `parse_model_action_line` converts that line into a `ModelAction`.
4. `validate_input_action` checks the placeholder and action.
5. `operate_mouse_keyboard` is the only place that should eventually call real OS mouse/keyboard APIs. It is dry-run until Tyler enables execution.

So the actual action source is not a separate script. It is currently the C function `local_model_stream_dry_run`; replace that function with the chosen local runtime bridge when moving from nub to live model.

## What “control console” means here

The control console is not ASCII control numbers 1-31. In this scaffold, “control console” means the C loop that builds a model-facing frame every second: screenshot path, mouse x/y, visible button rectangles, placeholder names, and overlay updates while a response continues.

Numbers mean:

- slots 1-40: training/input records in `training_slots.xml`;
- level 41: Markdown expansion instructions;
- `BTN_01`, `BTN_02`: placeholders for visible controls;
- `GRID_A1`, `GRID_B1`: optional screenshot grid regions.

## Getting user files into the scaffold

If GitHub will not upload documents, the clean path is to place files in a local folder next to the scaffold and point `training_slots.xml` at them. For chat-only handoff, split large content into chunks and paste them with stable filenames, for example `FILE: exhibit_a.txt PART 1/4`, then save each chunk into the target folder before listing it in XML.
