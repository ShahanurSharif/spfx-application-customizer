import {
  BaseApplicationCustomizer,
  PlaceholderName
} from '@microsoft/sp-application-base';
import styles from './HelloWorldCustomizer.module.scss';
import { NavMenuClickHandler } from './NavMenuClickHandler';
const editIcon = '<svg style="width: 16px; height: 16px;" viewBox="0 0 24 24"><path fill="currentColor" d="M20.71,7.04C21.1,6.65 21.1,6 20.71,5.63L18.37,3.29C18,2.9 17.35,2.9 16.96,3.29L15.12,5.12L18.87,8.87M3,17.25V21H6.75L17.81,9.93L14.06,6.18L3,17.25Z"></path></svg>';
/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IHelloWorldApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class HelloWorldApplicationCustomizer
  extends BaseApplicationCustomizer<IHelloWorldApplicationCustomizerProperties> {

  public onInit(): Promise<void> {
    this.context.placeholderProvider.changedEvent.add(
      this,
      this.top
    );
    this.context.placeholderProvider.changedEvent.add(
      this,
      this.bottom
    );
    
    return Promise.resolve();
  }

private bottom(): void {
    const placeholder = this.context.placeholderProvider.tryCreateContent(
      PlaceholderName.Bottom
    );
    if (placeholder) {
      placeholder.domElement.innerHTML = `
        <div class="${styles.siteTitle}">
          <h2>${this.properties.testMessage}</h2>
        </div>`;
    }
  }
  private top(): void {
    const placeholder = this.context.placeholderProvider.tryCreateContent(
      PlaceholderName.Top
    );
    if (placeholder) {
      setTimeout(() => {
        const navBar = document.querySelector('.ms-HorizontalNavItems-list');
        if (navBar) {
          const children = navBar.children;
          let homeItem: HTMLElement | undefined = undefined;
          for (let i = 0; i < children.length; i++) {
            const el = children[i] as HTMLElement;
            if (el.textContent && el.textContent.trim() === 'Pages') {
              homeItem = el;
              break;
            }
          }
          if (homeItem) {
            const iconButton = document.createElement('button');
            iconButton.title = 'Custom Action';
            iconButton.style.background = 'none';
            iconButton.style.border = 'none';
            iconButton.style.cursor = 'pointer';
            iconButton.style.fontSize = '1.3em';
            iconButton.style.marginLeft = '5px';
            iconButton.innerHTML = editIcon;
            // Add click event handler using the new class
            iconButton.addEventListener('click', function(event) {
              NavMenuClickHandler.handle(event, {
                elements: homeItem,
                id: homeItem ? homeItem.id || '' : '',
                index: Array.prototype.indexOf.call(children, homeItem),
                text: homeItem && homeItem.textContent ? homeItem.textContent.trim() : ''
              });
            });
            homeItem.appendChild(iconButton);
          }
        }
      }, 500);
      placeholder.domElement.innerHTML = `
        <div class="${styles.siteTitle}">
          <h2>Hello, World!</h2>
        </div>`;
    }
  }
}