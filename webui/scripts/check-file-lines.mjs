import { readdirSync, readFileSync, statSync } from 'node:fs'
import { join, relative } from 'node:path'

const root = process.cwd()
const sourceRoot = join(root, 'src')
const maxLines = Number.parseInt(process.env.SMARTOFFICE_MAX_SOURCE_LINES ?? '800', 10)
const extensions = new Set(['.ts', '.vue'])

function extensionOf(path) {
  const lastDot = path.lastIndexOf('.')
  return lastDot >= 0 ? path.slice(lastDot) : ''
}

function collectFiles(directory) {
  const files = []
  for (const entry of readdirSync(directory)) {
    if (entry === 'node_modules' || entry === 'dist') continue
    const fullPath = join(directory, entry)
    const stat = statSync(fullPath)
    if (stat.isDirectory()) {
      files.push(...collectFiles(fullPath))
    } else if (extensions.has(extensionOf(fullPath))) {
      files.push(fullPath)
    }
  }
  return files
}

function lineCount(path) {
  const text = readFileSync(path, 'utf8')
  if (!text) return 0
  return text.endsWith('\n') ? text.split('\n').length - 1 : text.split('\n').length
}

const oversized = collectFiles(sourceRoot)
  .map((path) => ({ path, lines: lineCount(path) }))
  .filter((item) => item.lines > maxLines)
  .sort((left, right) => right.lines - left.lines)

if (oversized.length > 0) {
  console.error(`Source file line-count gate failed. Max allowed lines: ${maxLines}.`)
  for (const item of oversized) {
    console.error(`${relative(root, item.path)}: ${item.lines}`)
  }
  process.exit(1)
}

console.log(`Source file line-count gate passed. Max allowed lines: ${maxLines}.`)
