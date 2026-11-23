<script setup lang="ts">
import { computed, ref, watch } from 'vue'
import { useRouter } from 'vue-router'
import { motion } from 'motion-v'
import { ArrowRight, Check, Download, FileSpreadsheet, Layers, LoaderCircle, Presentation, TriangleAlert } from 'lucide-vue-next'
import { Button } from '@/components/ui/button'
import DropZone from '@/components/DropZone.vue'
import { api, type Step, type SystemStatus } from '@/api'

const props = defineProps<{ status?: SystemStatus | null }>()
const router = useRouter()

// without LibreOffice (the hosted website) the deck comes as PowerPoint only; staff save it as PDF themselves
const steps = computed<{ key: Step; title: string; from: string; to: string; icon: any }[]>(() => {
  const pdf = props.status?.libreoffice !== false
  return [
    { key: 'portfolio', title: 'Portfolio Excel', from: 'Holdings template', to: 'Portfolio Excel', icon: FileSpreadsheet },
    { key: 'deck', title: 'Review deck', from: 'Portfolio Excel or template', to: pdf ? 'PowerPoint and PDF' : 'PowerPoint', icon: Presentation },
    { key: 'both', title: 'Excel and deck', from: 'Holdings template', to: pdf ? 'Excel, PowerPoint and PDF' : 'Excel and PowerPoint', icon: Layers },
  ]
})
const step = ref<Step>('both')
const file = ref<File | null>(null)
const error = ref('')
const busy = ref(false)

const hint = computed(() => step.value === 'deck'
  ? 'A Portfolio Excel made in step A (edited by hand is fine), or a filled holdings template.'
  : 'A filled CREST holdings template: client name in C2, as-on date in C3, holdings from row 6.')

const warnings = computed(() => {
  const s = props.status
  if (!s) return []
  const w: string[] = []
  if (!s.google) w.push('The Universal Sheet is read from the local file, not Google Sheets.')
  if (!s.ai) w.push('No narrative API key: ratings and the narrative will not be written by the AI.')
  else if (!s.web_search) w.push('No web search key: only blank sectors are filled; results stay “Results awaited”.')
  if (step.value !== 'portfolio' && !s.deck_template) w.push('The deck template is missing (templates/Deck_-_template.pptx).')
  if (!s.libreoffice && !s.hosted) w.push('LibreOffice is not installed: no PDF and no formula check.')
  return w
})

watch([file, step], () => { error.value = '' })

async function start() {
  if (!file.value) { error.value = 'Choose a file first.'; return }
  busy.value = true
  error.value = ''
  try {
    const job = await api.start(file.value, step.value)
    router.push({ name: 'run', params: { id: job.id } })
  } catch (e) {
    error.value = (e as Error).message
    busy.value = false
  }
}
</script>

<template>
  <div class="grid gap-10 lg:grid-cols-[minmax(0,1fr)_300px]">
    <section>
      <p class="text-sm font-medium tracking-wide text-primary">Quarterly portfolio review</p>
      <h1 class="mt-1 text-4xl">Start a new run</h1>
      <p class="mt-2 max-w-xl text-muted-foreground">
        Upload a client's holdings. Prices, results, commentary and ratings come live from the Universal Sheet.
      </p>

      <h2 class="mt-9 mb-3 font-sans text-sm font-medium text-muted-foreground">1 · What do you need?</h2>
      <div class="grid gap-3 sm:grid-cols-3">
        <motion.button
          v-for="(s, i) in steps"
          :key="s.key"
          type="button"
          class="relative rounded-xl border bg-card p-4 text-left transition-colors"
          :class="step === s.key ? 'border-primary ring-2 ring-primary/15' : 'hover:border-primary/40'"
          :initial="{ opacity: 0, y: 12 }"
          :animate="{ opacity: 1, y: 0 }"
          :transition="{ delay: 0.05 * i, duration: 0.25 }"
          :while-hover="{ y: -2 }"
          :while-press="{ scale: 0.98 }"
          @click="step = s.key"
        >
          <span class="mb-3 flex items-center justify-between">
            <span class="grid size-9 place-items-center rounded-lg" :class="step === s.key ? 'bg-primary text-primary-foreground' : 'bg-secondary text-primary'">
              <component :is="s.icon" class="size-4.5" />
            </span>
            <motion.span
              v-if="step === s.key"
              layout-id="step-check"
              class="grid size-5 place-items-center rounded-full bg-primary text-primary-foreground"
            >
              <Check class="size-3" />
            </motion.span>
          </span>
          <span class="block font-medium">{{ s.title }}</span>
          <span class="mt-1 flex items-center gap-1 text-xs text-muted-foreground">
            {{ s.from }} <ArrowRight class="size-3 shrink-0" /> {{ s.to }}
          </span>
        </motion.button>
      </div>

      <div class="mt-8 mb-3 flex items-end justify-between">
        <h2 class="font-sans text-sm font-medium text-muted-foreground">2 · Upload the file</h2>
        <a :href="api.templateUrl" class="inline-flex items-center gap-1.5 text-sm text-primary hover:underline">
          <Download class="size-3.5" /> Blank holdings template
        </a>
      </div>
      <DropZone v-model="file" :hint="hint" :invalid="!!error && !file" @reject="m => (error = m)" />

      <div class="mt-6 flex flex-wrap items-center gap-4">
        <Button size="lg" class="h-11 gap-2 px-6" :disabled="busy" @click="start">
          <LoaderCircle v-if="busy" class="size-4 animate-spin" />
          <template v-else>{{ { portfolio: 'Make the Portfolio Excel', deck: 'Make the deck', both: 'Make the Excel and deck' }[step] }} <ArrowRight class="size-4" /></template>
        </Button>
        <motion.p
          v-if="error"
          class="text-sm text-destructive"
          :initial="{ opacity: 0, x: -6 }"
          :animate="{ opacity: 1, x: 0 }"
        >{{ error }}</motion.p>
      </div>
    </section>

    <aside class="space-y-4 lg:pt-24">
      <div class="rounded-xl border bg-card p-5">
        <h3 class="font-sans text-sm font-medium">What happens</h3>
        <ol class="mt-3 space-y-3 text-sm text-muted-foreground">
          <li class="flex gap-3"><span class="step-dot">1</span>The holdings are checked; any problem names the exact cell.</li>
          <li class="flex gap-3"><span class="step-dot">2</span>Each stock is matched to the live Universal Sheet; gaps are researched by the AI.</li>
          <li class="flex gap-3"><span class="step-dot">3</span>The Excel and deck are built and checked. Every AI-written value is listed for review.</li>
        </ol>
      </div>
      <motion.div
        v-if="warnings.length"
        class="rounded-xl border border-warning/25 bg-warning-soft p-5 text-sm text-warning"
        :initial="{ opacity: 0 }"
        :animate="{ opacity: 1 }"
      >
        <p class="flex items-center gap-2 font-medium"><TriangleAlert class="size-4" /> Before you run</p>
        <ul class="mt-2 list-disc space-y-1 pl-5">
          <li v-for="w in warnings" :key="w">{{ w }}</li>
        </ul>
      </motion.div>
    </aside>
  </div>
</template>

<style scoped>
@reference "../style.css";
.step-dot {
  @apply grid size-5 shrink-0 place-items-center rounded-full bg-secondary text-[11px] font-medium text-primary;
}
</style>
