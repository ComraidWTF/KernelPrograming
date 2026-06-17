import {
  Component,
  Injectable,
  computed,
  effect,
  inject,
  input,
  signal,
} from '@angular/core';
import {
  GridModule,
  DataStateChangeEvent,
  ColumnReorderEvent,
  ColumnResizeArgs,
  ColumnVisibilityChangeEvent,
} from '@progress/kendo-angular-grid';
import {
  process,
  State,
  DataResult,
  CompositeFilterDescriptor,
} from '@progress/kendo-data-query';

/* ------------------------------------------------------------------ */
/*  Replace this with your real row type                               */
/* ------------------------------------------------------------------ */
export interface MyRow {
  id: number;
  name: string;
  email: string;
  createdAt: Date;
}

/* ------------------------------------------------------------------ */
/*  Persisted shapes                                                   */
/* ------------------------------------------------------------------ */
export interface ColumnSettings {
  field: string;
  title: string;
  width?: number;
  hidden?: boolean;
  orderIndex?: number;
  filter?: 'boolean' | 'numeric' | 'text' | 'date';
  filterable?: boolean;
}

export interface GridSettings {
  columns: ColumnSettings[];
  state: State;
}

/* ------------------------------------------------------------------ */
/*  Persistence service (localStorage)                                 */
/* ------------------------------------------------------------------ */
@Injectable({ providedIn: 'root' })
export class GridSettingsService {
  read(key: string): GridSettings | null {
    const raw = localStorage.getItem(key);
    if (!raw) return null;
    try {
      const parsed = JSON.parse(raw) as GridSettings;
      // JSON turns Date filter values into strings — revive them so
      // Kendo date filtering keeps working after a reload.
      reviveFilterDates(parsed.state?.filter);
      return parsed;
    } catch {
      return null;
    }
  }

  write(key: string, settings: GridSettings): void {
    localStorage.setItem(key, JSON.stringify(settings));
  }

  clear(key: string): void {
    localStorage.removeItem(key);
  }
}

function reviveFilterDates(filter?: CompositeFilterDescriptor): void {
  if (!filter) return;
  for (const f of filter.filters) {
    if ('filters' in f) {
      reviveFilterDates(f);
    } else if (
      typeof f.value === 'string' &&
      /^\d{4}-\d{2}-\d{2}T/.test(f.value)
    ) {
      f.value = new Date(f.value);
    }
  }
}

/* ------------------------------------------------------------------ */
/*  Defaults                                                           */
/* ------------------------------------------------------------------ */
const STORAGE_KEY = 'my-grid-settings';

const DEFAULT_COLUMNS: ColumnSettings[] = [
  { field: 'id', title: 'ID', width: 80, orderIndex: 0 },
  { field: 'name', title: 'Name', width: 200, orderIndex: 1 },
  { field: 'email', title: 'Email', width: 250, orderIndex: 2 },
  { field: 'createdAt', title: 'Created', width: 160, orderIndex: 3, filter: 'date' },
];

const DEFAULT_STATE: State = { skip: 0, take: 20 };

/* ------------------------------------------------------------------ */
/*  Component                                                          */
/* ------------------------------------------------------------------ */
@Component({
  selector: 'app-my-grid',
  standalone: true,
  imports: [GridModule],
  template: `
    <kendo-grid
      [data]="gridData()"
      [skip]="state().skip"
      [pageSize]="state().take"
      [sort]="state().sort"
      [filter]="state().filter"
      [group]="state().group"
      [pageable]="true"
      [sortable]="true"
      [filterable]="true"
      [groupable]="true"
      [reorderable]="true"
      [resizable]="true"
      (dataStateChange)="onStateChange($event)"
      (columnReorder)="onColumnReorder($event)"
      (columnResize)="onColumnResize($event)"
      (columnVisibilityChange)="onVisibilityChange($event)"
    >
      <ng-template kendoGridToolbarTemplate>
        <button kendoButton (click)="resetSettings()">Reset</button>
      </ng-template>

      @for (col of visibleColumns(); track col.field) {
        <kendo-grid-column
          [field]="col.field"
          [title]="col.title"
          [width]="col.width"
          [filter]="col.filter"
        ></kendo-grid-column>
      }
    </kendo-grid>
  `,
})
export class MyGridComponent {
  /** Your existing signal input — unchanged. */
  data = input.required<MyRow[]>();

  private readonly settings = inject(GridSettingsService);
  private readonly saved = this.settings.read(STORAGE_KEY);

  columns = signal<ColumnSettings[]>(
    this.saved?.columns ?? structuredClone(DEFAULT_COLUMNS)
  );
  state = signal<State>(this.saved?.state ?? { ...DEFAULT_STATE });

  /** Reactive processed data — recomputes when data() OR state() changes. */
  gridData = computed<DataResult>(() => process(this.data(), this.state()));

  /** What the template renders: visible columns in saved order. */
  visibleColumns = computed(() =>
    this.columns()
      .filter((c) => !c.hidden)
      .sort((a, b) => (a.orderIndex ?? 0) - (b.orderIndex ?? 0))
  );

  constructor() {
    // Persist whenever columns or state change. Note: data is NEVER
    // persisted — it always comes live from the input signal.
    effect(() => {
      this.settings.write(STORAGE_KEY, {
        columns: this.columns(),
        state: this.state(),
      });
    });
  }

  onStateChange(e: DataStateChangeEvent): void {
    this.state.set(e);
  }

  onColumnResize(e: ColumnResizeArgs[]): void {
    this.columns.update((cols) => {
      const next = [...cols];
      for (const r of e) {
        const field = (r.column as { field?: string }).field;
        const i = next.findIndex((c) => c.field === field);
        if (i > -1) next[i] = { ...next[i], width: r.newWidth };
      }
      return next;
    });
  }

  onColumnReorder(e: ColumnReorderEvent): void {
    const field = (e.column as { field?: string }).field;
    if (!field) return;
    this.columns.update((cols) => {
      const ordered = [...cols].sort(
        (a, b) => (a.orderIndex ?? 0) - (b.orderIndex ?? 0)
      );
      const from = ordered.findIndex((c) => c.field === field);
      if (from === -1) return cols;
      const [moved] = ordered.splice(from, 1);
      // newIndex is the target visible position; clamp to be safe.
      const to = Math.max(0, Math.min(e.newIndex, ordered.length));
      ordered.splice(to, 0, moved);
      return ordered.map((c, idx) => ({ ...c, orderIndex: idx }));
    });
  }

  onVisibilityChange(e: ColumnVisibilityChangeEvent): void {
    this.columns.update((cols) => {
      const next = [...cols];
      for (const col of e.columns) {
        const field = (col as { field?: string }).field;
        const i = next.findIndex((c) => c.field === field);
        if (i > -1) next[i] = { ...next[i], hidden: !!(col as { hidden?: boolean }).hidden };
      }
      return next;
    });
  }

  resetSettings(): void {
    this.settings.clear(STORAGE_KEY);
    this.columns.set(structuredClone(DEFAULT_COLUMNS));
    this.state.set({ ...DEFAULT_STATE });
  }
}
