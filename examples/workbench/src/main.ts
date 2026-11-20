import { createApp } from 'vue'
import 'antdv-next/dist/reset.css'
import 'antdv-next/dist/antd.css'
import '@office-studio/office-studio/style.css'
import App from './App.vue'
import { studio } from './runtime'

const app = createApp(App)
app.use(studio.pinia)
app.mount('#app')
