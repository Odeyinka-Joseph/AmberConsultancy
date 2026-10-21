import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
// base './' + HashRouter works on GitHub Pages under /<repo>/ with no 404 fallback needed.
export default defineConfig({ base: './', plugins: [react()] });
