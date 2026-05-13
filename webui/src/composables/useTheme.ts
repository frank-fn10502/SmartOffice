import { computed, ref, watch } from 'vue'

export type AppTheme = 'light' | 'dark'

const storageKey = 'smartoffice.theme'
const defaultTheme: AppTheme = 'dark'
const currentTheme = ref<AppTheme>(readInitialTheme())

function readInitialTheme(): AppTheme {
  if (typeof window === 'undefined') return defaultTheme
  const stored = window.localStorage.getItem(storageKey)
  return stored === 'light' || stored === 'dark' ? stored : defaultTheme
}

function applyTheme(theme: AppTheme) {
  if (typeof document === 'undefined') return
  const root = document.documentElement
  root.dataset.theme = theme
  root.dataset.colorMode = theme
  root.dataset.lightTheme = 'smartoffice-light'
  root.dataset.darkTheme = 'smartoffice-dark'
  root.classList.toggle('dark', theme === 'dark')
  root.style.colorScheme = theme
}

applyTheme(currentTheme.value)

watch(currentTheme, (theme) => {
  applyTheme(theme)
  if (typeof window !== 'undefined') {
    window.localStorage.setItem(storageKey, theme)
  }
})

export function useTheme() {
  const isDarkTheme = computed(() => currentTheme.value === 'dark')
  const themeLabel = computed(() => (isDarkTheme.value ? 'Dark' : 'Light'))

  function setTheme(theme: AppTheme) {
    currentTheme.value = theme
  }

  function toggleTheme() {
    setTheme(isDarkTheme.value ? 'light' : 'dark')
  }

  return {
    currentTheme,
    isDarkTheme,
    setTheme,
    themeLabel,
    toggleTheme,
  }
}
