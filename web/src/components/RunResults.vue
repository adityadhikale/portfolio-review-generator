<script setup lang="ts">
import { computed, ref } from 'vue'
import { motion } from 'motion-v'
import { Bot, ChevronRight, Download, ExternalLink, FileSpreadsheet, FileText, Presentation, ScrollText, Sparkles } from 'lucide-vue-next'
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs'
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table'
import StatusPill from '@/components/StatusPill.vue'
import { api, formatSize, type Job, type RunResult } from '@/api'

const props = defineProps<{ job: Job; result: RunResult }>()

const fileIcon = { excel: FileSpreadsheet, deck: Presentation, pdf: FileText, report: ScrollText, ai_ratings: Bot } as const
const banner = computed(() => {
  const s = props.result.status
  if (s === 'PASSED') return { cls: 'border-success/20 bg-success-soft', title: 'Ready to send', text: 'All checks passed.' }
  if (s === 'PASSED WITH WARNINGS') return { cls: 'border-warning/25 bg-warning-soft', title: 'Done, with points to review',
    text: 'The files are built. Look at the warnings and the AI-written values before sending.' }
  return { cls: 'border-destructive/20 bg-danger-soft', title: 'Stopped', text: stopped.value || 'A check failed; see below.' }
})
const stopped = computed(() => props.result.notes.find(n => n.startsWith('STOPPED'))?.replace(/^STOPPED \([^)]*\): /, ''))
const counts = computed(() => {
  const c = { PASS: 0, WARN: 0, FAIL: 0 }
  props.result.checks.forEach(x => (c[x.status] += 1))
  return c
})
const groups = computed(() => {
  const g: Record<string, typeof props.result.checks> = {}
  props.result.checks.forEach(c => (g[c.step || 'Checks'] ||= []).push(c))
  return g
})
const sevFilter = ref<'all' | 'FAIL' | 'WARN' | 'INFO'>('all')
const exceptions = computed(() => props.result.exceptions.filter(e => sevFilter.value === 'all' || e.severity === sevFilter.value))
const reviewCount = computed(() => props.result.ai_sourced.length)
const tab = ref(reviewCount.value ? 'review' : 'checks')
const links = (s: string) => s.split('|').map(x => x.trim()).filter(x => /^https?:\/\//.test(x))
const host = (u: string) => { try { return new URL(u).hostname.replace(/^www\./, '') } catch { return u } }
const metrics = computed(() => [
  { label: 'Holdings matched', value: `${props.result.matched}/${props.result.holdings ?? props.result.matched}` },
  { label: 'Checks passed', value: `${counts.value.PASS}/${props.result.checks.length}` },
  { label: 'AI values to review', value: String(reviewCount.value) },
  { label: 'Exceptions', value: String(props.result.exceptions.length) },
])
</script>

<template>
  <div class="space-y-6">
    <motion.div
      class="rounded-2xl border p-6"
      :class="banner.cls"
      :initial="{ opacity: 0, y: 10, scale: 0.99 }"
      :animate="{ opacity: 1, y: 0, scale: 1 }"
      :transition="{ duration: 0.3 }"
    >
      <div class="flex flex-wrap items-start justify-between gap-4">
        <div class="max-w-xl">
          <StatusPill :status="result.status" size="md" />
          <h2 class="mt-3 text-2xl">{{ banner.title }}</h2>
          <p class="mt-1 text-sm text-muted-foreground">{{ banner.text }}</p>
        </div>
        <dl class="grid grid-cols-[auto_auto] gap-x-4 gap-y-1 text-sm">
          <dt class="text-muted-foreground">Period</dt><dd class="font-medium">{{ result.period }}</dd>
          <dt class="text-muted-foreground">Holdings as on</dt><dd class="font-medium">{{ result.asof }}</dd>
          <dt class="text-muted-foreground">From</dt><dd class="font-medium">{{ result.holdings_kind }}</dd>
        </dl>
      </div>
    </motion.div>

    <div class="grid grid-cols-2 gap-3 md:grid-cols-4">
      <motion.div
        v-for="(m, i) in metrics"
        :key="m.label"
        class="rounded-xl border bg-card p-4"
        :initial="{ opacity: 0, y: 10 }"
        :animate="{ opacity: 1, y: 0 }"
        :transition="{ delay: 0.08 + i * 0.05 }"
      >
        <p class="text-xs text-muted-foreground">{{ m.label }}</p>
        <p class="mt-1 text-2xl font-medium tabular-nums">{{ m.value }}</p>
      </motion.div>
    </div>

    <div v-if="result.files.length" class="grid gap-3 sm:grid-cols-2 lg:grid-cols-3">
      <motion.a
        v-for="(f, i) in result.files"
        :key="f.name"
        :href="api.fileUrl(job.id, f.name)"
        class="group flex items-center gap-3 rounded-xl border bg-card p-4 transition-colors hover:border-primary/40"
        :initial="{ opacity: 0, y: 10 }"
        :animate="{ opacity: 1, y: 0 }"
        :transition="{ delay: 0.15 + i * 0.05 }"
        :while-hover="{ y: -2 }"
      >
        <span class="grid size-10 shrink-0 place-items-center rounded-lg bg-secondary text-primary">
          <component :is="fileIcon[f.kind] || FileText" class="size-5" />
        </span>
        <span class="min-w-0 flex-1">
          <span class="block text-sm font-medium">{{ f.label }}</span>
          <span class="block truncate text-xs text-muted-foreground">{{ f.name }} · {{ formatSize(f.size) }}</span>
        </span>
        <Download class="size-4 text-muted-foreground transition-colors group-hover:text-primary" />
      </motion.a>
    </div>

    <Tabs v-model="tab" class="mt-2">
      <TabsList class="h-10">
        <TabsTrigger value="review" class="gap-1.5 px-3">
          <Sparkles class="size-3.5" /> AI values to review
          <span v-if="reviewCount" class="ml-1 rounded-full bg-warning-soft px-1.5 text-[11px] text-warning">{{ reviewCount }}</span>
        </TabsTrigger>
        <TabsTrigger value="checks" class="px-3">Checks</TabsTrigger>
        <TabsTrigger value="exceptions" class="px-3">
          Exceptions <span class="ml-1 text-muted-foreground">{{ result.exceptions.length }}</span>
        </TabsTrigger>
        <TabsTrigger value="notes" class="px-3">Notes</TabsTrigger>
      </TabsList>

      <TabsContent value="review" class="mt-4">
        <div v-if="!reviewCount" class="rounded-xl border bg-card p-8 text-center text-sm text-muted-foreground">
          No AI-written Sector, Result Line or Commentary in this portfolio.
        </div>
        <div v-else class="rounded-xl border bg-card">
          <p class="border-b px-4 py-3 text-sm text-muted-foreground">
            Written by the AI, not typed by CREST. Check each against its source before the deck goes out. To correct one,
            type over it in the Universal Sheet and clear that row's Data Source.
          </p>
          <Table>
            <TableHeader>
              <TableRow>
                <TableHead class="w-32">Stock</TableHead>
                <TableHead class="w-36">Field</TableHead>
                <TableHead>Value</TableHead>
                <TableHead class="w-44">Sources</TableHead>
              </TableRow>
            </TableHeader>
            <TableBody>
              <TableRow v-for="r in result.ai_sourced" :key="r.symbol + r.field">
                <TableCell class="align-top font-medium">
                  {{ r.symbol }}
                  <span v-if="r.new" class="mt-1 block w-fit rounded bg-info-soft px-1.5 text-[11px] text-info">new</span>
                </TableCell>
                <TableCell class="align-top text-muted-foreground">{{ r.field }}</TableCell>
                <TableCell class="whitespace-normal align-top">{{ r.value }}</TableCell>
                <TableCell class="align-top">
                  <a v-for="u in links(r.sources)" :key="u" :href="u" target="_blank" rel="noopener"
                     class="flex items-center gap-1 truncate text-xs text-primary hover:underline">
                    <ExternalLink class="size-3 shrink-0" />{{ host(u) }}
                  </a>
                  <span v-if="!links(r.sources).length" class="text-xs text-muted-foreground">none recorded</span>
                </TableCell>
              </TableRow>
            </TableBody>
          </Table>
        </div>
        <div v-if="result.ai_research.skipped.length" class="mt-4 rounded-xl border bg-card p-4">
          <p class="text-sm font-medium">Not filled by the AI this run</p>
          <ul class="mt-2 space-y-1 text-sm text-muted-foreground">
            <li v-for="s in result.ai_research.skipped" :key="s.symbol + s.reason">
              <span class="font-medium text-foreground">{{ s.symbol }}</span> · {{ s.reason }}
            </li>
          </ul>
        </div>
      </TabsContent>

      <TabsContent value="checks" class="mt-4 space-y-4">
        <div v-for="(list, step) in groups" :key="step" class="rounded-xl border bg-card">
          <p class="border-b px-4 py-2.5 text-xs font-medium uppercase tracking-wider text-muted-foreground">{{ step }}</p>
          <details v-for="c in list" :key="c.name" class="group border-b last:border-0">
            <summary class="flex cursor-pointer list-none items-start gap-3 px-4 py-3 hover:bg-muted/50">
              <StatusPill :status="c.status" class="mt-0.5 w-24 justify-center" />
              <span class="min-w-0 flex-1">
                <span class="block text-sm font-medium">{{ c.name }}</span>
                <span class="block text-sm text-muted-foreground">{{ c.detail }}</span>
              </span>
              <ChevronRight v-if="c.items.length" class="mt-1 size-4 text-muted-foreground transition-transform group-open:rotate-90" />
            </summary>
            <ul v-if="c.items.length" class="space-y-1 bg-muted/40 px-4 py-3 pl-[8.5rem] text-sm text-muted-foreground">
              <li v-for="it in c.items" :key="it">{{ it }}</li>
            </ul>
          </details>
        </div>
      </TabsContent>

      <TabsContent value="exceptions" class="mt-4">
        <div class="mb-3 flex gap-1.5">
          <button v-for="f in (['all', 'FAIL', 'WARN', 'INFO'] as const)" :key="f" type="button"
                  class="rounded-full border px-3 py-1 text-xs transition-colors"
                  :class="sevFilter === f ? 'border-primary bg-primary text-primary-foreground' : 'bg-card hover:bg-muted'"
                  @click="sevFilter = f">
            {{ f === 'all' ? 'All' : f === 'FAIL' ? 'Failures' : f === 'WARN' ? 'Warnings' : 'Information' }}
          </button>
        </div>
        <div class="rounded-xl border bg-card">
          <Table>
            <TableHeader>
              <TableRow>
                <TableHead class="w-28">Severity</TableHead>
                <TableHead class="w-44">Type</TableHead>
                <TableHead class="w-32">Stock</TableHead>
                <TableHead>Detail</TableHead>
              </TableRow>
            </TableHeader>
            <TableBody>
              <TableRow v-for="(e, i) in exceptions" :key="i">
                <TableCell><StatusPill :status="e.severity" /></TableCell>
                <TableCell>{{ e.kind }}</TableCell>
                <TableCell class="font-medium">{{ e.symbol }}</TableCell>
                <TableCell class="whitespace-normal text-muted-foreground">{{ e.detail }}</TableCell>
              </TableRow>
              <TableRow v-if="!exceptions.length">
                <TableCell colspan="4" class="py-8 text-center text-muted-foreground">Nothing here.</TableCell>
              </TableRow>
            </TableBody>
          </Table>
        </div>
      </TabsContent>

      <TabsContent value="notes" class="mt-4">
        <ul class="divide-y rounded-xl border bg-card">
          <li v-for="n in result.notes" :key="n" class="px-4 py-3 text-sm"
              :class="n.startsWith('WARN') || n.startsWith('STOPPED') ? 'text-warning' : 'text-muted-foreground'">
            {{ n.replace(/^WARN: /, '') }}
          </li>
        </ul>
        <p class="mt-3 text-xs text-muted-foreground">
          Universal Sheet: {{ result.universal }}, read {{ result.universal_read_at }}.
          AI research: {{ result.ai_research.requests }} request(s), {{ result.ai_research.searches }} web search(es),
          {{ result.ai_research.filled }} value(s) filled.
        </p>
      </TabsContent>
    </Tabs>
  </div>
</template>
