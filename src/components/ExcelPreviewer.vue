<script setup lang="ts">
import sampleXlsxUrl from '@/assets/samples/sample.xlsx?url'
import type ExcelJS from 'exceljs'
import { computed, ref, shallowRef } from 'vue'

interface DisplayCell {
  text: string
  bgColor?: string
  fontColor?: string
  isBold?: boolean
  align?: 'left' | 'center' | 'right'
  rowspan: number
  colspan: number
  isHidden?: boolean
}

interface DisplayRow {
  number: number
  cells: DisplayCell[]
}

const workbookSheets = shallowRef<ExcelJS.Worksheet[]>([])
const selectedSheetName = ref('')
const fileName = ref('')
const errorMessage = ref('')
const isLoading = ref(false)

const activeSheet = computed(
  () => workbookSheets.value.find((sheet) => sheet.name === selectedSheetName.value) ?? null,
)

function parseColor(colorObj: any): string {
  if (!colorObj?.argb) return ''
  const argb = colorObj.argb
  return `#${argb.length === 8 ? argb.slice(2) : argb}`
}

const rows = computed<DisplayRow[]>(() => {
  const sheet = activeSheet.value
  if (!sheet) return []

  const maxRow = sheet.rowCount
  const maxCol = sheet.columnCount
  if (maxRow === 0 || maxCol === 0) return []

  const grid: DisplayCell[][] = []
  for (let r = 1; r <= maxRow; r++) {
    const rowCells: DisplayCell[] = []
    for (let c = 1; c <= maxCol; c++) {
      const cell = sheet.getCell(r, c)
      let bgColor = ''
      if (cell.fill && typeof cell.fill === 'object' && 'fgColor' in cell.fill) {
        bgColor = parseColor((cell.fill as any).fgColor)
      }

      const fontColor = cell.font ? parseColor(cell.font.color) : ''
      const isBold = !!cell.font?.bold
      let align: 'left' | 'center' | 'right' = 'left'
      const horizontal = cell.alignment?.horizontal
      if (horizontal === 'center' || horizontal === 'right' || horizontal === 'left') {
        align = horizontal
      }

      rowCells.push({
        text: formatCellValue(cell.value),
        bgColor,
        fontColor,
        isBold,
        align,
        rowspan: 1,
        colspan: 1,
        isHidden: false,
      })
    }
    grid.push(rowCells)
  }

  const merges = (sheet as any)._merges || (sheet as any).merges || []
  for (const rangeStr of Object.values(merges)) {
    let startRow = 0
    let startCol = 0
    let endRow = 0
    let endCol = 0

    if (typeof rangeStr === 'string') {
      const match = rangeStr.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/)
      if (match) {
        startCol = colLetterToNumber(match[1]!)
        startRow = parseInt(match[2]!, 10)
        endCol = colLetterToNumber(match[3]!)
        endRow = parseInt(match[4]!, 10)
      }
    } else if (rangeStr && typeof rangeStr === 'object') {
      startRow = (rangeStr as any).top || (rangeStr as any).model?.top
      startCol = (rangeStr as any).left || (rangeStr as any).model?.left
      endRow = (rangeStr as any).bottom || (rangeStr as any).model?.bottom
      endCol = (rangeStr as any).right || (rangeStr as any).model?.right
    }

    if (startRow && startCol && endRow && endCol) {
      const topLeft = grid[startRow - 1]?.[startCol - 1]
      if (!topLeft) continue
      topLeft.rowspan = endRow - startRow + 1
      topLeft.colspan = endCol - startCol + 1

      for (let r = startRow; r <= endRow; r++) {
        for (let c = startCol; c <= endCol; c++) {
          if (r === startRow && c === startCol) continue
          const coveredCell = grid[r - 1]?.[c - 1]
          if (coveredCell) coveredCell.isHidden = true
        }
      }
    }
  }

  return grid.map((cells, index) => ({ number: index + 1, cells }))
})

function colLetterToNumber(letter: string): number {
  let column = 0
  for (let i = 0; i < letter.length; i++) {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, letter.length - i - 1)
  }
  return column
}

const columns = computed(() =>
  Array.from({ length: activeSheet.value?.columnCount ?? 0 }, (_, index) => columnLabel(index + 1)),
)

function columnLabel(columnNumber: number): string {
  let number = columnNumber
  let label = ''

  while (number > 0) {
    const remainder = (number - 1) % 26
    label = String.fromCharCode(65 + remainder) + label
    number = Math.floor((number - 1) / 26)
  }

  return label
}

function formatCellValue(value: ExcelJS.CellValue): string {
  if (value === null || value === undefined) return ''
  if (value instanceof Date) return value.toLocaleString()
  if (typeof value !== 'object') return String(value)
  if ('error' in value) return value.error
  if ('richText' in value) return value.richText.map((part) => part.text).join('')
  if ('text' in value) return value.text
  if ('result' in value) {
    if (value.result === undefined) {
      const formula = 'formula' in value ? value.formula : value.sharedFormula
      return formula ? `=${formula}` : ''
    }
    return formatCellValue(value.result)
  }

  return ''
}

async function handleFileChange(event: Event) {
  const input = event.target
  if (!(input instanceof HTMLInputElement)) return

  const file = input.files?.[0]
  input.value = ''
  if (!file) return

  errorMessage.value = ''
  if (!/\.(xlsx|xlsm)$/i.test(file.name)) {
    errorMessage.value = '請選擇 .xlsx 或 .xlsm 格式的 Excel 檔案。'
    return
  }

  isLoading.value = true
  try {
    const { default: ExcelJS } = await import('exceljs')
    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.load(await file.arrayBuffer())

    if (workbook.worksheets.length === 0) {
      throw new Error('這個檔案沒有可顯示的工作表。')
    }

    workbookSheets.value = workbook.worksheets
    selectedSheetName.value = workbook.worksheets[0]!.name
    fileName.value = file.name
  } catch (error) {
    errorMessage.value =
      error instanceof Error
        ? `無法讀取這個 Excel 檔案：${error.message}`
        : '無法讀取這個 Excel 檔案，請確認檔案未損毀後再試一次。'
  } finally {
    isLoading.value = false
  }
}

// 載入 Sample 檔案的函式
async function loadSampleFile() {
  errorMessage.value = ''
  isLoading.value = true
  try {
    const response = await fetch(sampleXlsxUrl)
    if (!response.ok) throw new Error('無法取得 Sample 檔案')
    const arrayBuffer = await response.arrayBuffer()

    const { default: ExcelJS } = await import('exceljs')
    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.load(arrayBuffer)

    if (workbook.worksheets.length === 0) {
      throw new Error('Sample 檔案中沒有可顯示的工作表。')
    }

    workbookSheets.value = workbook.worksheets
    selectedSheetName.value = workbook.worksheets[0]!.name
    fileName.value = 'sample.xlsx'
  } catch (error) {
    errorMessage.value =
      error instanceof Error
        ? `無法載入 Sample 檔案：${error.message}`
        : '無法載入 Sample 檔案，請確認檔案路徑是否正確。'
  } finally {
    isLoading.value = false
  }
}
</script>

<template>
  <section class="w-full min-w-0 text-[#172b24]">
    <header
      class="mb-6 flex flex-col items-start justify-between gap-6 sm:mb-8 sm:flex-row sm:items-center"
    >
      <div>
        <h1 class="text-[clamp(1.7rem,3vw,2.2rem)] font-bold leading-tight tracking-[-0.04em]">
          Excel 檔案預覽
        </h1>
      </div>
    </header>

    <p
      v-if="errorMessage"
      class="-mt-3 mb-5 rounded-[0.65rem] border border-[#f2c5c2] bg-[#fff5f4] px-4 py-3 text-sm text-[#a32a22]"
      role="alert"
    >
      {{ errorMessage }}
    </p>

    <!-- 主區塊：設定固定高度 h-[720px]，使用 flex 垂直排列 -->
    <section
      v-if="activeSheet"
      class="flex h-[720px] w-full flex-col overflow-hidden rounded-[0.9rem] border border-[#e4eae6] bg-white shadow-[0_12px_35px_#1933280a]"
      aria-label="Excel 工作表內容"
    >
      <!-- 檔案資訊列 (固定高度) -->
      <div
        class="flex min-h-[4.2rem] shrink-0 items-center justify-between gap-4 border-b border-[#e4eae6] px-4 py-3 sm:px-[1.1rem]"
      >
        <div class="flex min-w-0 items-center gap-3">
          <span
            class="grid size-[2.3rem] shrink-0 place-items-center rounded-[0.45rem] bg-[#e9f5ee] text-[0.68rem] font-extrabold text-[#16845b]"
            aria-hidden="true"
          >
            XLS
          </span>
          <div class="grid min-w-0 gap-px">
            <strong class="truncate text-sm font-semibold text-[#172b24]">{{ fileName }}</strong>
            <span class="text-[0.78rem] text-[#718078]">{{ workbookSheets.length }} 個工作表</span>
          </div>
        </div>
        <div class="hidden items-center gap-3 text-[0.78rem] text-[#718078] sm:flex">
          <span>列數：{{ activeSheet.rowCount }}</span>
          <span>•</span>
          <span>欄數：{{ activeSheet.columnCount }}</span>
        </div>
      </div>

      <!-- 表格檢視區：使用 flex-1 填滿剩餘高度，overflow-auto 支援雙向 (上下與左右) 捲動 -->
      <div v-if="rows.length && columns.length" class="flex-1 min-h-0 overflow-auto">
        <table
          class="w-max min-w-full border-separate border-spacing-0 text-left text-[0.84rem] text-[#25332d]"
        >
          <thead>
            <tr>
              <th
                class="sticky top-0 left-0 z-30 min-w-14 border-r border-b border-[#e9eeeb] bg-[#eef2ef] px-3 py-2 text-center text-xs font-semibold text-[#53635b]"
                aria-label="列號"
              ></th>
              <th
                v-for="column in columns"
                :key="column"
                class="sticky top-0 z-20 min-w-22 border-r border-b border-[#e9eeeb] bg-[#f4f7f5] px-3 py-2 text-center text-xs font-semibold text-[#53635b]"
              >
                {{ column }}
              </th>
            </tr>
          </thead>
          <tbody>
            <template v-for="row in rows" :key="row.number">
              <tr class="hover:[&>td]:brightness-[0.97]">
                <th
                  class="sticky left-0 z-10 w-14 min-w-14 max-w-14 border-r border-b border-[#e9eeeb] bg-[#f8faf9] px-3 py-2 text-center text-xs font-medium text-[#87938c]"
                  scope="row"
                >
                  {{ row.number }}
                </th>
                <template v-for="(cell, index) in row.cells">
                  <!-- 略過被合併儲存格涵蓋的隱藏儲存格 -->
                  <td
                    v-if="!cell.isHidden"
                    :key="index"
                    :rowspan="cell.rowspan > 1 ? cell.rowspan : undefined"
                    :colspan="cell.colspan > 1 ? cell.colspan : undefined"
                    class="h-[2.45rem] min-w-34 max-w-96 wrap-anywhere border-r border-b border-[#e9eeeb] px-3 py-[0.45rem] whitespace-pre-wrap"
                    :style="{
                      backgroundColor: cell.bgColor || '',
                      color: cell.fontColor || '',
                      fontWeight: cell.isBold ? 'bold' : 'normal',
                      textAlign: cell.align,
                    }"
                  >
                    {{ cell.text }}
                  </td>
                </template>
              </tr>
            </template>
          </tbody>
        </table>
      </div>

      <div
        v-else
        class="flex flex-1 flex-col items-center justify-center gap-1 border-t border-[#e4eae6] px-4 py-12 text-center text-[0.86rem] text-[#718078]"
      >
        <span class="mb-1 text-3xl text-[#9aa8a0]" aria-hidden="true">▦</span>
        <strong class="font-semibold text-[#172b24]">這個工作表目前是空的</strong>
        <span>選擇其他工作表，或上傳另一個檔案。</span>
      </div>

      <!-- Excel 風格底部工作表頁籤列 -->
      <div
        class="flex shrink-0 items-center gap-1 overflow-x-auto border-t border-[#e4eae6] bg-[#f0f4f1] px-2 py-1.5"
        role="tablist"
        aria-label="工作表分頁"
      >
        <button
          v-for="sheet in workbookSheets"
          :key="sheet.id"
          role="tab"
          :aria-selected="sheet.name === selectedSheetName"
          class="flex items-center gap-2 rounded-t-md border-b-2 px-4 py-2 text-xs font-medium transition-colors whitespace-nowrap"
          :class="[
            sheet.name === selectedSheetName
              ? 'border-[#16845b] bg-white text-[#16845b] shadow-sm font-semibold'
              : 'border-transparent text-[#53635b] hover:bg-[#e2ebe5] hover:text-[#172b24]',
          ]"
          @click="selectedSheetName = sheet.name"
        >
          <span
            class="size-2 rounded-full"
            :class="sheet.name === selectedSheetName ? 'bg-[#16845b]' : 'bg-[#b0bdb6]'"
          ></span>
          {{ sheet.name }}
        </button>
      </div>
    </section>

    <!-- 尚未上傳時的佔位畫面 (加入載入 Sample 按鈕) -->
    <section
      v-else-if="!errorMessage"
      class="flex min-h-80 flex-col items-center justify-center gap-4 rounded-[0.9rem] border border-dashed border-[#cbd8d0] bg-[#fbfdfb] p-6 text-center sm:min-h-96 sm:p-8"
    >
      <div class="flex flex-wrap items-center justify-center gap-3">
        <label
          class="inline-flex min-h-10 cursor-pointer items-center justify-center rounded-[0.65rem] border border-[#16845b] bg-[#16845b] px-4 py-2 text-sm font-semibold text-white transition-colors hover:border-[#116b49] hover:bg-[#116b49] focus-within:outline-3 focus-within:outline-offset-2 focus-within:outline-[#16845b44]"
        >
          選擇本地檔案
          <input
            type="file"
            accept=".xlsx,.xlsm"
            :disabled="isLoading"
            class="sr-only"
            @change="handleFileChange"
          />
        </label>

        <button
          type="button"
          :disabled="isLoading"
          class="inline-flex min-h-10 cursor-pointer items-center justify-center rounded-[0.65rem] border border-[#cbd8d0] bg-white px-4 py-2 text-sm font-semibold text-[#172b24] transition-colors hover:bg-[#f4f7f5] focus-within:outline-3 focus-within:outline-offset-2 focus-within:outline-[#16845b44]"
          @click="loadSampleFile"
        >
          載入 Sample 檔案
        </button>
      </div>
    </section>
  </section>
</template>
