<script setup lang="ts">
import { computed, onMounted, onUnmounted, ref } from 'vue'
import { RouterLink, useRouter } from 'vue-router'
import { motion } from 'motion-v'
import { ArrowRight, FileSpreadsheet, Plus, Search } from 'lucide-vue-next'
import { Button } from '@/components/ui/button'
import { Skeleton } from '@/components/ui/skeleton'
import StatusPill from '@/components/StatusPill.vue'
import { api, formatWhen, type HistoryRow } from '@/api'

const router = useRouter()
const rows = ref<HistoryRow[] | null>(null)
const error = ref('')
const q = ref('')
let timer: number | undefined

async function load() {
  try { rows.value = await api.history(); error.value = '' } catch (e) { error.value = (e as Error).message }
}
onMounted(() => { load(); timer = window.setInterval(load, 5000) })
onUnmounted(() => window.clearInterval(timer))

const shown = computed(() => (rows.value || []).filter(r =>
  !q.value || `${r.client ?? ''} ${r.filename} ${r.period ?? ''}`.toLowerCase().includes(q.value.toLowerCase())))
const label = (r: HistoryRow) => (r.status === 'done' ? r.run_status : r.status)
</script>

<template>
  <div>
    <div class="flex flex-wrap items-end justify-between gap-4">
      <div>
        <h1 class="text-4xl">History</h1>
        <p class="mt-2 text-muted-foreground">Every run, with its files. Newest first.</p>
      </div>
      <label class="relative">
        <Search class="pointer-events-none absolute top-1/2 left-3 size-4 -translate-y-1/2 text-muted-foreground" />
        <input v-model="q" placeholder="Client, file or period" class="h-10 w-64 rounded-lg border bg-card pr-3 pl-9 text-sm outline-none focus:ring-2 focus:ring-ring/30" />
      </label>
    </div>

    <p v-if="error" class="mt-6 text-sm text-destructive">{{ error }}</p>

    <div v-if="rows === null" class="mt-8 space-y-2">
      <Skeleton v-for="i in 5" :key="i" class="h-16 w-full rounded-xl" />
    </div>

    <div v-else-if="!rows.length" class="mt-10 rounded-2xl border bg-card p-12 text-center">
      <span class="mx-auto grid size-12 place-items-center rounded-full bg-secondary text-primary"><FileSpreadsheet class="size-6" /></span>
      <p class="mt-4 font-medium">Start your first run</p>
      <p class="mt-1 text-sm text-muted-foreground">Upload a client's holdings to make the Portfolio Excel and the review deck.</p>
      <RouterLink to="/"><Button class="mt-5 gap-2"><Plus class="size-4" /> New run</Button></RouterLink>
    </div>

    <div v-else class="mt-8 overflow-hidden rounded-xl border bg-card">
      <motion.button
        v-for="(r, i) in shown"
        :key="r.id"
        type="button"
        class="group grid w-full grid-cols-[minmax(0,1.6fr)_minmax(0,1fr)_minmax(0,1fr)_auto_auto] items-center gap-4 border-b px-5 py-4 text-left last:border-0 hover:bg-muted/50"
        :initial="{ opacity: 0, y: 6 }"
        :animate="{ opacity: 1, y: 0 }"
        :transition="{ delay: Math.min(i, 12) * 0.03 }"
        @click="router.push({ name: 'run', params: { id: r.id } })"
      >
        <span class="min-w-0">
          <span class="block truncate font-medium">{{ r.client || r.filename }}</span>
          <span class="block truncate text-xs text-muted-foreground">{{ r.filename }}</span>
        </span>
        <span class="text-sm text-muted-foreground">{{ r.step_label }}<span v-if="r.period"> · {{ r.period }}</span></span>
        <span class="text-sm text-muted-foreground">{{ formatWhen(r.created) }}</span>
        <StatusPill :status="label(r)" />
        <ArrowRight class="size-4 text-muted-foreground transition-transform group-hover:translate-x-0.5" />
      </motion.button>
      <p v-if="!shown.length" class="p-8 text-center text-sm text-muted-foreground">No run matches “{{ q }}”.</p>
    </div>
  </div>
</template>
