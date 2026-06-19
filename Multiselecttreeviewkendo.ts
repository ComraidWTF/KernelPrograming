import { Component } from '@angular/core';

type TreeFilterItem = {
  id: string;
  text: string;
  items?: TreeFilterItem[];
};

@Component({
  selector: 'app-tree-filter',
  template: `
    <kendo-multiselecttree
      [data]="filters"
      textField="text"
      valueField="id"
      childrenField="items"
      [valuePrimitive]="true"
      [filterable]="true"
      [(ngModel)]="selectedIds"
      [isItemDisabled]="isParentNode">
    </kendo-multiselecttree>

    <pre>{{ selectedIds | json }}</pre>
  `
})
export class TreeFilterComponent {
  selectedIds: string[] = [];

  filters: TreeFilterItem[] = [
    {
      id: 'location',
      text: 'Location',
      items: [
        {
          id: 'uk',
          text: 'UK',
          items: [
            { id: 'scotland', text: 'Scotland' },
            { id: 'england', text: 'England' }
          ]
        },
        {
          id: 'india',
          text: 'India',
          items: [
            { id: 'kerala', text: 'Kerala' },
            { id: 'karnataka', text: 'Karnataka' }
          ]
        }
      ]
    },
    {
      id: 'status',
      text: 'Status',
      items: [
        { id: 'active', text: 'Active' },
        { id: 'inactive', text: 'Inactive' }
      ]
    },
    {
      id: 'standalone',
      text: 'Standalone Item'
    }
  ];

  isParentNode = (item: TreeFilterItem): boolean => {
    return !!item.items?.length;
  };
}
