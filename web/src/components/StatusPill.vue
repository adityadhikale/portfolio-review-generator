<script setup lang="ts">
import { computed } from 'vue'
import { CircleAlert, CircleCheck, CircleX, Clock, LoaderCircle } from 'lucide-vue-next'

const props = defineProps<{ status: string | null | undefined; size?: 'sm' | 'md' }>()

const look = computed(() => {
  const s = (props.status || '').toUpperCase()
  if (s === 'PASS' || s === 'PASSED' || s === 'DONE') return { cls: 'bg-success-soft text-success', icon: CircleCheck, label: s === 'PASS' ? 'Pass' : 'Passed' }
  if (s === 'WARN' || s === 'PASSED WITH WARNINGS') return { cls: 'bg-warning-soft text-warning', icon: CircleAlert, label: s === 'WARN' ? 'Warning' : 'Passed with warnings' }
  if (s === 'FAIL' || s === 'FAILED') return { cls: 'bg-danger-soft text-destructive', icon: CircleX, label: s === 'FAIL' ? 'Fail' : 'Failed' }
  if (s === 'RUNNING') return { cls: 'bg-info-soft text-info', icon: LoaderCircle, label: 'Running', spin: true }
  if (s === 'QUEUED') return { cls: 'bg-muted text-muted-foreground', icon: Clock, label: 'Queued' }
  if (s === 'INFO') return { cls: 'bg-info-soft text-info', icon: CircleAlert, label: 'Info' }
  return { cls: 'bg-muted text-muted-foreground', icon: Clock, label: props.status || '' }
})
</script>

<template>
  <span
    class="inline-flex items-center gap-1 whitespace-nowrap rounded-full font-medium"
    :class="[look.cls, size === 'md' ? 'px-3 py-1 text-sm' : 'px-2 py-0.5 text-xs']"
  >
    <component :is="look.icon" :class="[size === 'md' ? 'size-4' : 'size-3.5', look.spin ? 'animate-spin' : '']" />
    {{ look.label }}
  </span>
</template>
