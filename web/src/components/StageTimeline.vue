<script setup lang="ts">
import { computed, onMounted, onUnmounted, ref } from 'vue'
import { AnimatePresence, motion } from 'motion-v'
import { Check, LoaderCircle } from 'lucide-vue-next'
import type { Job } from '@/api'

const props = defineProps<{ job: Job }>()
const now = ref(Date.now())
let t: number | undefined
onMounted(() => { t = window.setInterval(() => (now.value = Date.now()), 1000) })
onUnmounted(() => window.clearInterval(t))

const current = computed(() => {
  const i = props.job.stages.findIndex(s => s.key === props.job.stage)
  return i < 0 ? 0 : i
})
const elapsed = computed(() => {
  if (!props.job.started) return '0:00'
  const s = Math.max(0, Math.floor((now.value - new Date(props.job.started).getTime()) / 1000))
  return `${Math.floor(s / 60)}:${String(s % 60).padStart(2, '0')}`
})
const pct = computed(() => Math.round(((current.value + 0.5) / props.job.stages.length) * 100))
</script>

<template>
  <div class="rounded-2xl border bg-card p-6 sm:p-8">
    <div class="flex items-baseline justify-between">
      <div>
        <p class="text-sm text-muted-foreground">Working on it</p>
        <p class="mt-0.5 text-lg font-medium">{{ job.stages[current]?.label }}…</p>
      </div>
      <p class="font-mono text-sm tabular-nums text-muted-foreground">{{ elapsed }}</p>
    </div>

    <div class="mt-5 h-1.5 overflow-hidden rounded-full bg-muted">
      <motion.div
        class="relative h-full rounded-full bg-primary"
        :animate="{ width: `${pct}%` }"
        :transition="{ type: 'spring', stiffness: 60, damping: 18 }"
      >
        <span class="shimmer absolute inset-0" />
      </motion.div>
    </div>

    <ol class="mt-7 space-y-1">
      <motion.li
        v-for="(s, i) in job.stages"
        :key="s.key"
        class="flex items-center gap-3 rounded-lg px-2 py-2"
        :class="i === current ? 'bg-accent' : ''"
        :initial="{ opacity: 0, x: -8 }"
        :animate="{ opacity: 1, x: 0 }"
        :transition="{ delay: 0.04 * i }"
      >
        <span class="relative grid size-6 shrink-0 place-items-center">
          <AnimatePresence mode="wait">
            <motion.span
              v-if="i < current"
              key="done"
              class="grid size-6 place-items-center rounded-full bg-primary text-primary-foreground"
              :initial="{ scale: 0.4, opacity: 0 }"
              :animate="{ scale: 1, opacity: 1 }"
              :transition="{ type: 'spring', stiffness: 500, damping: 25 }"
            ><Check class="size-3.5" /></motion.span>
            <motion.span v-else-if="i === current" key="now" class="text-primary" :initial="{ opacity: 0 }" :animate="{ opacity: 1 }">
              <LoaderCircle class="size-5 animate-spin" />
            </motion.span>
            <span v-else key="todo" class="size-2 rounded-full bg-border" />
          </AnimatePresence>
        </span>
        <span class="text-sm" :class="i < current ? 'text-foreground' : i === current ? 'font-medium text-accent-foreground' : 'text-muted-foreground'">
          {{ s.label }}
        </span>
      </motion.li>
    </ol>
    <p class="mt-6 text-xs text-muted-foreground">
      A run takes one to three minutes; the AI and the PDF are the slow parts. You can leave this page; the run carries on
      and stays in History.
    </p>
  </div>
</template>

<style scoped>
.shimmer {
  background: linear-gradient(90deg, transparent, rgba(255, 255, 255, 0.45), transparent);
  animation: shimmer 1.6s infinite;
}
@keyframes shimmer {
  from { transform: translateX(-100%); }
  to { transform: translateX(100%); }
}
</style>
