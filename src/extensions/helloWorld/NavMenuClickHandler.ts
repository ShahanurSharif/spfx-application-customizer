import { Dialog } from '@microsoft/sp-dialog';

export interface INavMenuClickEvent {
  elements: HTMLElement | null;
  id: string;
  index: number;
  text: string;
}

export class NavMenuClickHandler {
  public static handle(event: MouseEvent, menuItem: INavMenuClickEvent): void {
    Dialog.alert(
      `Menu item clicked!\nElement: ${JSON.stringify(menuItem.elements)}\n ${menuItem.id}\nIndex: ${menuItem.index}\nText: ${menuItem.text}`
    ).catch(() => {/* ignore */});
  }
}
