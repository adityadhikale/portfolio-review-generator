<script setup lang="ts">
import { onMounted, onUnmounted, ref } from 'vue'
import { RouterLink, RouterView } from 'vue-router'
import { AnimatePresence, motion } from 'motion-v'
import { FileSpreadsheet, History, Plus } from 'lucide-vue-next'
import { api, type SystemStatus } from '@/api'
import StatusChips from '@/components/StatusChips.vue'

const status = ref<SystemStatus | null>(null)
let timer: number | undefined

async function refresh() {
  try { status.value = await api.status() } catch { status.value = null }
}
onMounted(() => { refresh(); timer = window.setInterval(refresh, 30000) })
onUnmounted(() => window.clearInterval(timer))
</script>

<template>
  <div class="min-h-screen">
    <header class="sticky top-0 z-30 border-b bg-background/85 backdrop-blur">
      <div class="mx-auto flex h-16 max-w-6xl items-center gap-6 px-6">
        <RouterLink to="/" class="flex items-center gap-2.5">
          <span class="grid size-8 place-items-center rounded-lg bg-primary text-primary-foreground">
            <FileSpreadsheet class="size-4.5" />
          </span>
          <span class="leading-tight">
            <span class="block font-display text-[17px] tracking-wide">CREST</span>
            <span class="block text-[11px] uppercase tracking-[0.14em] text-muted-foreground">Portfolio Review</span>
          </span>
        </RouterLink>

        <nav class="flex items-center gap-1 text-sm">
          <RouterLink to="/" class="nav-link" exact-active-class="nav-active">
            <Plus class="size-4" /> New run
          </RouterLink>
          <RouterLink to="/history" class="nav-link" active-class="nav-active">
            <History class="size-4" /> History
          </RouterLink>
        </nav>

        <StatusChips class="ml-auto" :status="status" />
      </div>
    </header>

    <main class="mx-auto max-w-6xl px-6 py-10">
      <RouterView v-slot="{ Component, route }">
        <AnimatePresence mode="wait">
          <motion.div
            :key="route.fullPath"
            :initial="{ opacity: 0, y: 10 }"
            :animate="{ opacity: 1, y: 0 }"
            :exit="{ opacity: 0, y: -6 }"
            :transition="{ duration: 0.22, ease: 'easeOut' }"
          >
            <component :is="Component" :status="status" />
          </motion.div>
        </AnimatePresence>
      </RouterView>
    </main>
  </div>
</template>

<style scoped>
@reference "./style.css";
.nav-link {
  @apply inline-flex items-center gap-1.5 rounded-md px-3 py-1.5 text-muted-foreground transition-colors hover:bg-muted hover:text-foreground;
}
.nav-active {
  @apply bg-secondary text-secondary-foreground;
}
</style>
