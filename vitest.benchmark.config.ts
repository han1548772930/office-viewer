import { defineConfig } from 'vitest/config'

export default defineConfig({
  test: {
    include: ['tests/benchmarks/**/*.smoke.test.mjs'],
    environment: 'node',
    pool: 'forks',
    maxWorkers: 1,
    testTimeout: 150_000,
    hookTimeout: 30_000,
  },
})
