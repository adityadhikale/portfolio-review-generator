<script setup lang="ts">
import { ref } from 'vue'
import { AnimatePresence, motion } from 'motion-v'
import { FileSpreadsheet, UploadCloud, X } from 'lucide-vue-next'
import { formatSize } from '@/api'

const props = defineProps<{ hint: string; invalid?: boolean }>()
const file = defineModel<File | null>({ default: null })
const emit = defineEmits<{ reject: [message: string] }>()
const over = ref(false)
const input = ref<HTMLInputElement>()

function take(f: File | undefined | null) {
  if (!f) return
  if (!/\.(xlsx|xlsm)$/i.test(f.name)) {
    emit('reject', 'Choose an Excel file (.xlsx): a filled holdings template or a Portfolio Excel.')
    return
  }
  file.value = f
}
function onDrop(e: DragEvent) {
  over.value = false
  take(e.dataTransfer?.files?.[0])
}
</script>

<template>
  <div
    role="button"
    tabindex="0"
    class="group relative flex min-h-48 cursor-pointer flex-col items-center justify-center overflow-hidden rounded-xl border-2 border-dashed bg-card px-6 py-8 text-center transition-colors"
    :class="[over ? 'border-primary bg-accent' : props.invalid ? 'border-destructive/60' : 'border-input hover:border-primary/50 hover:bg-accent/40']"
    @click="input?.click()"
    @keydown.enter.space.prevent="input?.click()"
    @dragover.prevent="over = true"
    @dragleave.prevent="over = false"
    @drop.prevent="onDrop"
  >
    <input ref="input" type="file" accept=".xlsx,.xlsm" class="hidden" @change="take(($event.target as HTMLInputElement).files?.[0]); ($event.target as HTMLInputElement).value = ''" />
    <AnimatePresence mode="wait">
      <motion.div
        v-if="!file"
        key="empty"
        class="flex flex-col items-center"
        :initial="{ opacity: 0, scale: 0.97 }"
        :animate="{ opacity: 1, scale: 1 }"
        :exit="{ opacity: 0, scale: 0.97 }"
        :transition="{ duration: 0.18 }"
      >
        <motion.div
          class="mb-3 grid size-12 place-items-center rounded-full bg-secondary text-primary"
          :animate="over ? { y: -4, scale: 1.08 } : { y: 0, scale: 1 }"
          :transition="{ type: 'spring', stiffness: 400, damping: 22 }"
        >
          <UploadCloud class="size-6" />
        </motion.div>
        <p class="font-medium">{{ over ? 'Drop it here' : 'Drag the file here, or click to choose' }}</p>
        <p class="mt-1 max-w-md text-sm text-muted-foreground">{{ hint }}</p>
      </motion.div>
      <motion.div
        v-else
        key="file"
        class="flex w-full max-w-md items-center gap-3 rounded-lg border bg-background p-3 text-left"
        :initial="{ opacity: 0, y: 8 }"
        :animate="{ opacity: 1, y: 0 }"
        :exit="{ opacity: 0, y: -8 }"
        :transition="{ duration: 0.2 }"
        @click.stop
      >
        <span class="grid size-10 shrink-0 place-items-center rounded-md bg-success-soft text-success">
          <FileSpreadsheet class="size-5" />
        </span>
        <span class="min-w-0 flex-1">
          <span class="block truncate font-medium">{{ file.name }}</span>
          <span class="text-xs text-muted-foreground">{{ formatSize(file.size) }} · ready to run</span>
        </span>
        <button
          type="button"
          class="rounded-md p-1.5 text-muted-foreground transition-colors hover:bg-muted hover:text-foreground"
          aria-label="Remove file"
          @click.stop="file = null"
        >
          <X class="size-4" />
        </button>
      </motion.div>
    </AnimatePresence>
  </div>
</template>
