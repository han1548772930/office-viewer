import vue from '@vitejs/plugin-vue'
import { defineConfig } from 'vite'

export default defineConfig({
  plugins: [vue()],
  build: {
    emptyOutDir: false,
    cssCodeSplit: false,
    lib: {
      entry: 'src/index.ts',
      formats: ['es'],
      fileName: 'office-studio',
      cssFileName: 'office-studio',
    },
    rollupOptions: {
      external: ['vue', 'pinia', 'antdv-next', 'lucide-vue-next', /^antdv-next\//, /^lucide-vue-next\//],
      output: { globals: { vue: 'Vue', pinia: 'Pinia' } },
    },
  },
})
