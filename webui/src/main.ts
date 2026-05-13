import { createApp } from 'vue'
import ElementPlus from 'element-plus'
import 'element-plus/dist/index.css'
import 'element-plus/theme-chalk/dark/css-vars.css'

import App from './App.vue'
import './styles.css'
import './features/outlook/styles/mail.css'

createApp(App).use(ElementPlus).mount('#app')
