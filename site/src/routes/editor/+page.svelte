<script lang="ts">
  import EditorApp from '$lib/editor/EditorApp.svelte';
</script>

<svelte:head>
  <title>xlsx editor — @office-kit/xlsx</title>
</svelte:head>

<div class="editor-page">
  <EditorApp />
</div>

<style>
  .editor-page {
    position: fixed;
    inset: 0;
  }
</style>
