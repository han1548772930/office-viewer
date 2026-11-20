import { createApp } from 'vue'
import { createPinia } from 'pinia'
import 'antdv-next/dist/reset.css'
import 'antdv-next/dist/antd.css'
import './style.css'
import App from './App.vue'

const app = createApp(App)
app.use(createPinia())
app.mount('#app')
