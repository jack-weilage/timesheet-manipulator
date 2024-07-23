<script lang="ts">
	import { invoke } from "@tauri-apps/api/tauri";

	let running = false;
	let error: string | undefined;

	async function begin_parse_worksheet() {
		// Don't attempt to parse a worksheet if another parse is already in progress.
		if (running) {
			return;
		}

		// Unset the error from the previous run.
		error = undefined;
		running = true;

		await invoke("parse_worksheet").catch((err) => (error = err));
		running = false;
	}
</script>

<main class="flex h-screen flex-col p-2">
	<label class="flex flex-col">
		Browse to the data file, then select the output location.
		<button
			type="button"
			class="my-2 rounded-md border bg-zinc-200 px-4 py-1 hover:bg-zinc-300"
			on:click={begin_parse_worksheet}
			disabled={running}
		>
			Browse
		</button>
	</label>
	{#if error}
		<p class="text-rose-500">An error occurred: {error}</p>
	{/if}
</main>
