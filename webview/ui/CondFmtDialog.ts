import { WebviewState } from '../state/WebviewState';
import { messageBridge } from '../state/MessageBridge';
import { getElement } from './domUtils';

export class CondFmtDialog {
  private dialog: HTMLElement;

  constructor(private state: WebviewState) {
    this.dialog = getElement('cond-fmt-dialog');
    this.setup();
  }

  open(): void {
    this.dialog.classList.add('open');
    this.renderRules();
    this.updateCondTypeVisibility();
  }

  private updateCondTypeVisibility(): void {
    const condType = getElement<HTMLSelectElement>('cond-type');
    const isScale = condType.value === 'colorScale';
    const isBar = condType.value === 'dataBar';
    const isBetween = condType.value === 'between';

    getElement('cond-value2-label').style.display = isBetween ? 'flex' : 'none';
    getElement('cond-value1').closest('label')!.style.display =
      isScale || isBar ? 'none' : '';
    getElement('cond-standard-style').style.display = isScale || isBar ? 'none' : '';
    getElement('cond-colorscale-opts').style.display = isScale ? '' : 'none';
    getElement('cond-databar-opts').style.display = isBar ? '' : 'none';
  }

  private renderRules(): void {
    const list = getElement('cond-rules-list');
    const sheet = this.state.activeSheet;
    list.innerHTML = '';
    if (!sheet || sheet.conditionalFormats.length === 0) return;

    sheet.conditionalFormats.forEach((rule) => {
      const div = document.createElement('div');
      div.className = 'cond-rule';

      const label = document.createElement('span');
      if (rule.condition === 'colorScale') {
        const cs = rule.colorScaleColors;
        label.innerHTML = `Color scale: <span style="display:inline-block;width:40px;height:10px;background:linear-gradient(to right,${cs?.min ?? '#ff0000'},${cs?.max ?? '#00ff00'});border:1px solid #666;vertical-align:middle"></span>`;
      } else if (rule.condition === 'dataBar') {
        label.innerHTML = `Data bar <span style="display:inline-block;width:12px;height:12px;background:${rule.dataBarColor ?? '#4472C4'};border:1px solid #666;vertical-align:middle;opacity:0.7"></span>`;
      } else {
        const swatch = rule.style.backgroundColor
          ? `<span style="display:inline-block;width:12px;height:12px;background:${rule.style.backgroundColor};border:1px solid #666;margin-right:4px;vertical-align:middle"></span>`
          : '';
        label.innerHTML = swatch + rule.condition + ' ' + rule.value1 + (rule.value2 ? ' - ' + rule.value2 : '');
      }
      div.appendChild(label);

      const removeBtn = document.createElement('span');
      removeBtn.className = 'remove-rule';
      removeBtn.textContent = '\u00d7';
      removeBtn.addEventListener('click', () => {
        messageBridge.postMessage({ type: 'removeConditionalFormat', sheet: sheet.name, ruleId: rule.id });
      });
      div.appendChild(removeBtn);
      list.appendChild(div);
    });
  }

  private setup(): void {
    const condType = getElement<HTMLSelectElement>('cond-type');
    condType.addEventListener('change', () => this.updateCondTypeVisibility());

    getElement('cond-apply').addEventListener('click', () => {
      const sheet = this.state.activeSheet;
      const range = this.state.selectionRange;
      if (!sheet || !range) return;

      const type = condType.value;

      if (type === 'colorScale') {
        const minColor = (getElement<HTMLInputElement>('cs-min-color')).value;
        const midColor = (getElement<HTMLInputElement>('cs-mid-color')).value;
        const maxColor = (getElement<HTMLInputElement>('cs-max-color')).value;
        const useMid = (getElement<HTMLInputElement>('cs-use-mid')).checked;
        messageBridge.postMessage({
          type: 'setConditionalFormat',
          sheet: sheet.name,
          rule: {
            id: 'cf-' + Date.now(),
            range,
            condition: 'colorScale',
            value1: '',
            style: {},
            colorScaleColors: { min: minColor, mid: useMid ? midColor : undefined, max: maxColor },
          },
        });
      } else if (type === 'dataBar') {
        const barColor = (getElement<HTMLInputElement>('db-color')).value;
        messageBridge.postMessage({
          type: 'setConditionalFormat',
          sheet: sheet.name,
          rule: {
            id: 'cf-' + Date.now(),
            range,
            condition: 'dataBar',
            value1: '',
            style: {},
            dataBarColor: barColor,
          },
        });
      } else {
        const value1 = (getElement<HTMLInputElement>('cond-value1')).value;
        const value2 = (getElement<HTMLInputElement>('cond-value2')).value;
        const bgColor = (getElement<HTMLInputElement>('cond-bg')).value;
        const textColor = (getElement<HTMLInputElement>('cond-text')).value;
        const bold = (getElement<HTMLInputElement>('cond-bold')).checked;
        const italic = (getElement<HTMLInputElement>('cond-italic')).checked;

        messageBridge.postMessage({
          type: 'setConditionalFormat',
          sheet: sheet.name,
          rule: {
            id: 'cf-' + Date.now(),
            range,
            condition: type,
            value1,
            value2: type === 'between' ? value2 : undefined,
            style: {
              backgroundColor: bgColor,
              textColor,
              ...(bold ? { bold: true } : {}),
              ...(italic ? { italic: true } : {}),
            },
          },
        });
      }

      this.dialog.classList.remove('open');
    });

    getElement('cond-cancel').addEventListener('click', () => {
      this.dialog.classList.remove('open');
    });
  }
}
