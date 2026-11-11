<script setup>
import { Download, Trash2 } from 'lucide-vue-next'

defineProps({
  ouvert: Boolean,
  historique: { type: Array, default: () => [] },
  chargement: Boolean
})
const emit = defineEmits(['telecharger', 'supprimer'])

function formaterDate(dateIso) {
  const d = new Date(dateIso.replace(' ', 'T') + 'Z')
  return d.toLocaleDateString('fr-FR', { day: '2-digit', month: 'short' }) +
    ', ' + d.toLocaleTimeString('fr-FR', { hour: '2-digit', minute: '2-digit' })
}

function formaterKo(octets) {
  return (octets / 1024).toFixed(0) + ' Ko'
}
</script>

<template>
  <aside
    :class="[
      'shrink-0 overflow-hidden border-r border-line bg-paper-dim transition-[width] duration-300 ease-in-out',
      ouvert ? 'w-72' : 'w-0 border-r-0'
    ]"
  >
    <div class="w-72 h-full flex flex-col">
      <div class="px-4 py-3.5 border-b border-line">
        <span class="font-display text-[13px] font-semibold text-ink">Fichiers convertis</span>
      </div>

      <div class="flex-1 overflow-y-auto">
        <div v-if="chargement" class="px-4 py-6 font-mono text-[11px] text-ink-soft">
          chargement…
        </div>

        <div v-else-if="!historique.length" class="px-4 py-8">
          <p class="text-xs text-ink-soft leading-relaxed">
            Rien pour l'instant. Un fichier converti apparaît ici, prêt à être retéléchargé.
          </p>
        </div>

        <ul v-else>
          <li
            v-for="entree in historique"
            :key="entree.id"
            class="group relative border-b border-line/60 pl-4 pr-2 py-3 hover:bg-paper transition-colors"
          >
            <div class="absolute left-0 top-0 bottom-0 w-0.5 bg-yellow scale-y-0 group-hover:scale-y-100 transition-transform origin-center" />

            <div class="flex items-start justify-between gap-2">
              <p class="min-w-0 truncate text-[13px] font-medium text-ink" :title="entree.nom_original">
                {{ entree.nom_original }}
              </p>
              <span class="shrink-0 font-mono text-[10px] text-ink-soft pt-0.5">{{ formaterKo(entree.taille_octets) }}</span>
            </div>

            <div class="mt-1 flex items-center justify-between gap-2">
              <span class="flex min-w-0 items-center gap-1.5 font-mono text-[10px] text-ink-soft">
                <span class="shrink-0 rounded-sm bg-blue-dim px-1 font-medium leading-4 text-blue">{{ entree.style_entete }}</span>
                <span class="truncate">
                  {{ formaterDate(entree.date_creation) }}<template v-if="entree.nb_feuilles > 1"> · {{ entree.nb_feuilles }} fichiers</template>
                </span>
              </span>

              <div class="flex shrink-0 items-center gap-0.5 opacity-0 group-hover:opacity-100 group-focus-within:opacity-100 transition-opacity">
                <button
                  type="button"
                  @click="emit('telecharger', entree)"
                  class="flex h-6 w-6 items-center justify-center rounded text-ink-soft hover:text-blue hover:bg-blue-dim"
                  aria-label="Télécharger"
                >
                  <Download :size="12" :stroke-width="2" />
                </button>
                <button
                  type="button"
                  @click="emit('supprimer', entree)"
                  class="flex h-6 w-6 items-center justify-center rounded text-ink-soft hover:text-danger hover:bg-danger/10"
                  aria-label="Supprimer"
                >
                  <Trash2 :size="12" :stroke-width="2" />
                </button>
              </div>
            </div>
          </li>
        </ul>
      </div>
    </div>
  </aside>
</template>
