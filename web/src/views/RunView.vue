<script setup lang="ts">
import { onUnmounted, ref, watch } from 'vue'
import { RouterLink } from 'vue-router'
import { AnimatePresence, motion } from 'motion-v'
import { ArrowLeft, Clock, RotateCcw, TriangleAlert } from 'lucide-vue-next'
import { Button } from '@/components/ui/button'
import { Skeleton } from '@/components/ui/skeleton'
import StageTimeline from '@/components/StageTimeline.vue'
import RunResults from '@/components/RunResults.vue'
import StatusPill from '@/components/StatusPill.vue'
import { api, formatWhen, type Job } from '@/api'

const props = defineProps<{ id: string }>()
const job = ref<Job | null>(null)
const missing = ref('')
let stop: (() => void) | undefined
let poll: number | undefined

async function load(id: string) {
  stop?.()
  window.clearInterval(poll)
  job.value = null
  missing.value = ''
  try {
    job.value = await api.job(id)
  } catch (e) {
    missing.value = (e as Error).message
    return
  }
  if (job.value.status === 'queued' || job.value.status === 'running') {
    stop = api.watch(id, j => (job.value = j), () => {
      // live updates dropped (server restarted, network): fall back to polling
      poll = window.setInterval(async () => {
        try {
          job.value = await api.job(id)
          if (job.value.status === 'done' || job.value.status === 'failed') window.clearInterval(poll)
        } catch { /* keep trying */ }
      }, 3000)
    })
  }
}
watch(() => props.id, load, { immediate: true })
onUnmounted(() => { stop?.(); window.clearInterval(poll) })
</script>

<template>
  <div>
    <RouterLink to="/history" class="inline-flex items-center gap-1.5 text-sm text-muted-foreground hover:text-foreground">
      <ArrowLeft class="size-4" /> All runs
    </RouterLink>

    <div v-if="missing" class="mt-8 rounded-xl border bg-card p-8 text-center">
      <p class="font-medium">This run could not be found.</p>
      <p class="mt-1 text-sm text-muted-foreground">{{ missing }}</p>
    </div>

    <div v-else-if="!job" class="mt-8 space-y-4">
      <Skeleton class="h-10 w-72" />
      <Skeleton class="h-64 w-full rounded-2xl" />
    </div>

    <template v-else>
      <div class="mt-4 mb-8 flex flex-wrap items-end justify-between gap-3">
        <div>
          <h1 class="text-3xl">{{ job.client || job.filename }}</h1>
          <p class="mt-1 text-sm text-muted-foreground">
            {{ job.step_label }} · {{ job.filename }} · started {{ formatWhen(job.created) }}
          </p>
        </div>
        <StatusPill v-if="job.status !== 'done'" :status="job.status" size="md" />
      </div>

      <AnimatePresence mode="wait">
        <motion.div
          v-if="job.status === 'queued'"
          key="queued"
          class="flex items-center gap-4 rounded-2xl border bg-card p-8"
          :initial="{ opacity: 0 }" :animate="{ opacity: 1 }" :exit="{ opacity: 0 }"
        >
          <motion.span
            class="grid size-12 place-items-center rounded-full bg-secondary text-primary"
            :animate="{ scale: [1, 1.08, 1] }"
            :transition="{ duration: 1.6, repeat: Infinity }"
          ><Clock class="size-6" /></motion.span>
          <div>
            <p class="font-medium">Waiting in the queue{{ job.queue_position ? ` · number ${job.queue_position}` : '' }}</p>
            <p class="text-sm text-muted-foreground">Runs go one at a time. This one starts as soon as the run ahead finishes.</p>
          </div>
        </motion.div>

        <motion.div v-else-if="job.status === 'running'" key="running"
                    :initial="{ opacity: 0 }" :animate="{ opacity: 1 }" :exit="{ opacity: 0 }">
          <StageTimeline :job="job" />
        </motion.div>

        <motion.div
          v-else-if="job.status === 'failed'"
          key="failed"
          class="rounded-2xl border border-destructive/20 bg-danger-soft p-8"
          :initial="{ opacity: 0, y: 10 }" :animate="{ opacity: 1, y: 0 }"
        >
          <p class="flex items-center gap-2 font-medium text-destructive"><TriangleAlert class="size-5" /> The run could not start</p>
          <p class="mt-3 whitespace-pre-line text-sm">{{ job.error }}</p>
          <RouterLink to="/"><Button variant="outline" class="mt-5 gap-2"><RotateCcw class="size-4" /> Fix the file and try again</Button></RouterLink>
        </motion.div>

        <motion.div v-else-if="job.result" key="done" :initial="{ opacity: 0 }" :animate="{ opacity: 1 }">
          <RunResults :job="job" :result="job.result" />
        </motion.div>
      </AnimatePresence>
    </template>
  </div>
</template>
