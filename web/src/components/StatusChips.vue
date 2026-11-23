<script setup lang="ts">
import { computed } from 'vue'
import { Bot, Globe, Search, Sheet, TriangleAlert } from 'lucide-vue-next'
import { Tooltip, TooltipContent, TooltipProvider, TooltipTrigger } from '@/components/ui/tooltip'
import type { SystemStatus } from '@/api'

const props = defineProps<{ status: SystemStatus | null }>()

const chips = computed(() => {
  const s = props.status
  if (!s) return []
  return [
    { icon: Sheet, ok: s.google, label: s.google ? 'Google Sheet' : 'Local sheet',
      tip: s.google ? 'Universal Sheet read live from Google Sheets' : 'Using the local Universal Sheet file (offline fallback)' },
    { icon: Bot, ok: s.ai, label: s.ai ? 'AI on' : 'AI off',
      tip: s.ai ? 'Ratings and narrative by the narrative API' : 'No narrative API key in .env: no AI ratings or narrative' },
    { icon: Search, ok: s.web_search, label: s.web_search ? 'Web search' : 'No search',
      tip: s.web_search ? 'Result lines and commentary researched with Tavily' : 'No Tavily key: only blank sectors are filled' },
    s.hosted && !s.libreoffice
      ? { icon: Globe, ok: true, label: 'PPT only',
          tip: 'Hosted website: the deck comes as PowerPoint. Open it and use File > Save as PDF.' }
      : { icon: Globe, ok: s.libreoffice, label: s.libreoffice ? 'PDF ready' : 'No PDF',
          tip: s.libreoffice ? 'LibreOffice found: formula checks and PDF' : 'LibreOffice not found: no PDF or formula check' },
  ]
})
</script>

<template>
  <TooltipProvider :delay-duration="150">
    <div class="flex items-center gap-1.5">
      <div v-if="!status" class="flex items-center gap-1.5 rounded-full bg-danger-soft px-3 py-1 text-xs text-destructive">
        <TriangleAlert class="size-3.5" /> Server not reachable
      </div>
      <Tooltip v-for="c in chips" :key="c.label">
        <TooltipTrigger as-child>
          <span
            class="inline-flex items-center gap-1.5 rounded-full border px-2.5 py-1 text-xs"
            :class="c.ok ? 'border-transparent bg-success-soft text-success' : 'border-border bg-muted text-muted-foreground'"
          >
            <component :is="c.icon" class="size-3.5" />
            <span class="hidden lg:inline">{{ c.label }}</span>
          </span>
        </TooltipTrigger>
        <TooltipContent>{{ c.tip }}</TooltipContent>
      </Tooltip>
    </div>
  </TooltipProvider>
</template>
