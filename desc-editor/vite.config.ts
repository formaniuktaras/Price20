import { defineConfig, type PluginOption } from 'vite';

const reactPluginPromise = import('@vitejs/plugin-react')
  .then(({ default: react }) => react())
  .catch(() => null);

export default defineConfig(async () => {
  const reactPlugin = await reactPluginPromise;
  const plugins: PluginOption[] = [];

  if (reactPlugin) {
    plugins.push(reactPlugin);
  }

  return {
    plugins,
    server: {
      port: 5173,
    },
  };
});
