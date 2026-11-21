import { createApp } from 'vue'
import 'antdv-next/dist/reset.css'
import 'antdv-next/dist/antd.css'
import '@office-studio/office-studio/style.css'
import './workspace-shell.css'
import App from './App.vue'
import { studio } from './runtime'
import { createOfficeI18n } from '@office-studio/office-studio'

const app = createApp(App)
app.use(studio.pinia)
app.use(createOfficeI18n())
app.mount('#app')

if (import.meta.hot) import.meta.hot.dispose(() => app.unmount())
