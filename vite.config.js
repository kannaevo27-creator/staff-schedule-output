import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// GitHub Pages のサブパスに合わせる:
// https://<user>.github.io/staff-schedule-output/
export default defineConfig({
  plugins: [react()],
  base: '/staff-schedule-output/',
})
