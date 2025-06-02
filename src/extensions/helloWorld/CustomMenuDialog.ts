import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';

export interface ICustomMenuDialogProps {
  id: string;
  index: number;
  text: string;
  onSubmit: (color: string, size: number) => void;
}

export class CustomMenuDialog extends BaseDialog {
  private props: ICustomMenuDialogProps;
  private color: string = '#0078d4';
  private size: number = 16;

  public constructor(props: ICustomMenuDialogProps) {
    super();
    this.props = props;
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div style="margin-bottom: 12px;">
        <div><strong>ID:</strong> ${this.props.id}</div>
        <div><strong>Index:</strong> ${this.props.index}</div>
        <div><strong>Text:</strong> ${this.props.text}</div>
      </div>
      <div style="margin-bottom: 12px;">
        <label for="colorSlider"><strong>Color:</strong></label>
        <input id="colorSlider" type="color" value="${this.color}" style="margin-left: 8px;" />
      </div>
      <div style="margin-bottom: 12px;">
        <label for="sizeSlider"><strong>Size:</strong></label>
        <input id="sizeSlider" type="range" min="10" max="50" value="${this.size}" style="margin-left: 8px;" />
        <span id="sizeValue">${this.size}</span>px
      </div>
      <button id="submitBtn" style="margin-top: 12px; padding: 6px 16px; background: #0078d4; color: white; border: none; border-radius: 3px; cursor: pointer;">Submit</button>
    `;

    const colorSlider = this.domElement.querySelector('#colorSlider') as HTMLInputElement;
    const sizeSlider = this.domElement.querySelector('#sizeSlider') as HTMLInputElement;
    const sizeValue = this.domElement.querySelector('#sizeValue') as HTMLSpanElement;
    const submitBtn = this.domElement.querySelector('#submitBtn') as HTMLButtonElement;

    if (colorSlider) {
      colorSlider.addEventListener('input', () => {
        this.color = colorSlider.value;
      });
    }
    if (sizeSlider && sizeValue) {
      sizeSlider.addEventListener('input', () => {
        this.size = Number(sizeSlider.value);
        sizeValue.textContent = sizeSlider.value;
      });
    }
    if (submitBtn) {
      submitBtn.addEventListener('click', () => {
        this.props.onSubmit(this.color, this.size);
        this.close();
      });
    }
  }

  public getConfig(): IDialogConfiguration {
    return {
      isBlocking: false
    };
  }
}
