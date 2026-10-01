import { BaseListViewCommandSet, type IListViewCommandSetExecuteEventParameters, type ListViewStateChangedEventArgs } from '@microsoft/sp-listview-extensibility';

const DORK_LIST_ID = 'dd8ac8ef-3be9-4fb8-b7cb-6ac0f56d2b03';
const GEEK_URL = 'https://thankful-sky-0cc1b0210.6.azurestaticapps.net/geek.html';

export default class GeekLauncherCommandSet extends BaseListViewCommandSet<Record<string, never>> {
  public onInit(): Promise<void> {
    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);
    this._updateVisibility();
    return Promise.resolve();
  }
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    if (event.itemId !== 'GEEK_UPLOAD' || !this._isDorkLibrary()) return;
    window.open(GEEK_URL, '_blank', 'noopener,noreferrer');
  }
  public onDispose(): void {
    this.context.listView.listViewStateChangedEvent.remove(this, this._onListViewStateChanged);
    super.onDispose();
  }
  private _isDorkLibrary(): boolean {
    return this.context.pageContext.list?.id.toString().toLowerCase() === DORK_LIST_ID;
  }
  private _updateVisibility(): void {
    const command = this.tryGetCommand('GEEK_UPLOAD');
    if (command) command.visible = this._isDorkLibrary() && (this.context.listView.selectedRows?.length || 0) === 0;
  }
  private _onListViewStateChanged = (_args: ListViewStateChangedEventArgs): void => {
    this._updateVisibility();
    this.raiseOnChange();
  };
}
