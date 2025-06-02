import { CustomMenuDialog } from './CustomMenuDialog';

export interface INavMenuClickEvent {
  elements: HTMLElement | undefined;
  id: string;
  index: number;
  text: string;
}

export class NavMenuClickHandler {
  public static handle(event: MouseEvent, menuItem: INavMenuClickEvent): void {
    // Show the custom interactive dialog
    const dialog = new CustomMenuDialog({
      id: menuItem.id,
      index: menuItem.index,
      text: menuItem.text,
      onSubmit: (color: string, size: number) => {
        // You can handle the submit event here (e.g., log, update UI, etc.)
        // eslint-disable-next-line no-console
        console.log('Submitted:', { id: menuItem.id, index: menuItem.index, text: menuItem.text, color, size });
      }
    });
    dialog.show().catch(() => {/* ignore */});
  }
}
