declare function acquireVsCodeApi(): {
  postMessage(msg: unknown): void;
  getState(): unknown;
  setState(state: unknown): void;
};

const vscode = acquireVsCodeApi();

type MessageHandler = (msg: unknown) => void;

class MessageBridgeImpl {
  private handlers: MessageHandler[] = [];

  constructor() {
    window.addEventListener('message', (event) => {
      for (const handler of this.handlers) {
        handler(event.data);
      }
    });
  }

  onMessage(handler: MessageHandler): void {
    this.handlers.push(handler);
  }

  postMessage(msg: unknown): void {
    vscode.postMessage(msg);
  }
}

export const messageBridge = new MessageBridgeImpl();
