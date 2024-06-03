import * as React from "react";
import { useState, useEffect } from "react";
import {
  DefaultButton,
  PrimaryButton,
} from "office-ui-fabric-react/lib/Button";
import { Panel, PanelType } from "office-ui-fabric-react/lib/Panel";
import {
  Dialog,
  DialogType,
  DialogFooter,
} from "office-ui-fabric-react/lib/Dialog";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { Spinner, SpinnerSize } from "office-ui-fabric-react/lib/Spinner";
import {
  MessageBar,
  MessageBarType,
} from "office-ui-fabric-react/lib/MessageBar";
import { List } from "office-ui-fabric-react/lib/List";
import {
  FocusZone,
  FocusZoneDirection,
} from "office-ui-fabric-react/lib/FocusZone";
import { IMyFavouritesTopBarProps } from "./IMyFavouritesTopBarProps";
import { MyFavouritesService } from "../../../services/MyFavouritesService";
import { IMyFavouriteItem } from "../../../interfaces/IMyFavouriteItem";
import MyFavouriteDisplayItem from "../MyFavouriteDisplayItem/MyFavouriteDisplayItem";
import { css } from "@uifabric/utilities/lib/css";
import styles from "../MyFavourites.module.scss";
import * as strings from "MyFavouritesApplicationCustomizerStrings";

let _MyFavouriteItems: IMyFavouriteItem[] = [];

const MyFavouritesTopBar: React.FC<IMyFavouritesTopBarProps> = (props) => {
  const [showPanel, setShowPanel] = useState(false);
  const [showDialog, setShowDialog] = useState(false);
  const [dialogTitle, setDialogTitle] = useState("");
  const [myFavouriteItems, setMyFavouriteItems] = useState<IMyFavouriteItem[]>(
    []
  );
  const [itemInContext, setItemInContext] = useState<IMyFavouriteItem>({
    Id: 0,
    Title: "",
    Description: "",
  });
  const [isEdit, setIsEdit] = useState(false);
  const [status, setStatus] = useState(
    <Spinner size={SpinnerSize.large} label={strings.LoadingStatusLabel} />
  );
  const [disableButtons, setDisableButtons] = useState(false);

  const _MyFavouritesServiceInstance = new MyFavouritesService(props);

  const editFavourite = (favouriteItem: IMyFavouriteItem): void => {
    let status: JSX.Element = <span></span>;
    let dialogTitle: string = strings.EditFavouritesDialogTitle;
    setShowPanel(false);
    setItemInContext(favouriteItem);
    setIsEdit(true);
    setShowDialog(true);
    setDialogTitle(dialogTitle);
    setStatus(status);
  };

  const _getMyFavourites = async (): Promise<void> => {
    let status: JSX.Element = (
      <Spinner size={SpinnerSize.large} label={strings.LoadingStatusLabel} />
    );
    setStatus(status);

    const myFavouriteItems: IMyFavouriteItem[] =
      await _MyFavouritesServiceInstance.getMyFavourites(true);
    _MyFavouriteItems = myFavouriteItems;
    status = <span></span>;
    setMyFavouriteItems(myFavouriteItems);
    setStatus(status);
  };

  const deleteFavourite = async (favouriteItemId: number): Promise<void> => {
    let result: boolean = await _MyFavouritesServiceInstance.deleteFavourite(
      favouriteItemId
    );
    if (result) {
      _getMyFavourites();
    }
  };

  const _hideMenu = () => {
    setShowPanel(false);
  };

  const _hideDialog = () => {
    setShowDialog(false);
  };

  const _saveMyFavourite = async (): Promise<void> => {
    let status: JSX.Element = (
      <Spinner size={SpinnerSize.large} label={strings.SavingStatusLabel} />
    );
    setDisableButtons(true);
    setStatus(status);
    let itemToSave: IMyFavouriteItem = {
      Title: itemInContext.Title,
      Description: itemInContext.Description,
    };
    let itemToEdit: IMyFavouriteItem = {
      ...itemToSave,
      Id: itemInContext.Id,
    };
    let result: boolean = isEdit
      ? await _MyFavouritesServiceInstance.updateFavourite(itemToEdit)
      : await _MyFavouritesServiceInstance.saveFavourite(itemToSave);

    if (result) {
      _hideDialog();
    } else {
      status = (
        <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
          There was an error!
        </MessageBar>
      );
    }
    setDisableButtons(false);
    setStatus(status);
  };

  const _showMenu = () => {
    _getMyFavourites();
    setShowPanel(true);
  };

  const _showDialog = () => {
    let itemInContext: IMyFavouriteItem = {
      Id: 0,
      Title: "",
      Description: "",
    };
    let isEdit: boolean = false;
    let status: JSX.Element = <span></span>;
    let dialogTitle: string = strings.AddToFavouritesDialogTitle;
    setItemInContext(itemInContext);
    setIsEdit(isEdit);
    setShowDialog(true);
    setDialogTitle(dialogTitle);
    setStatus(status);
  };

  const _onRenderCell = (
    myFavouriteItem: IMyFavouriteItem,
    _index: number | undefined
  ): JSX.Element => {
    return (
      <div
        className={css("ms-slideDownIn20", styles.ccitemCell)}
        data-is-focusable={true}
      >
        <MyFavouriteDisplayItem
          displayItem={myFavouriteItem}
          deleteFavourite={(favouriteItemId: number) =>
            deleteFavourite(favouriteItemId)
          }
          editFavoutite={editFavourite}
        />
      </div>
    );
  };

  const _onFilterChanged = (
    _event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ) => {
    let items: IMyFavouriteItem[] = _MyFavouriteItems;
    setMyFavouriteItems(
      newValue
        ? items.filter(
            (item) =>
              Number(
                item.Title?.toLowerCase().indexOf(newValue.toLowerCase())
              ) >= 0
          )
        : items
    );
  };

  const _setItemInContext = (
    _event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ) => {
    setItemInContext({
      ...itemInContext,
      [(_event.target as HTMLInputElement).name]: newValue,
    });
  };

  useEffect(() => {
    _getMyFavourites();
  }, []);

  return (
    <div className={styles.ccTopBar}>
      <PrimaryButton
        data-id="menuButton"
        title={strings.ShowMyFavouritesLabel}
        text={strings.ShowMyFavouritesLabel}
        ariaLabel={strings.ShowMyFavouritesLabel}
        iconProps={{ iconName: "View" }}
        className={styles.ccTopBarButton}
        onClick={_showMenu}
      />
      <PrimaryButton
        data-id="menuButton"
        title={strings.AddPageToFavouritesLabel}
        text={strings.AddPageToFavouritesLabel}
        ariaLabel={strings.AddPageToFavouritesLabel}
        iconProps={{ iconName: "Add" }}
        className={styles.ccTopBarButton}
        onClick={_showDialog}
      />
      <Panel
        isOpen={showPanel}
        type={PanelType.medium}
        onDismiss={_hideMenu}
        headerText={strings.MyFavouritesHeader}
        isLightDismiss={true}
      >
        <div data-id="menuPanel">
          <TextField
            placeholder={strings.FilterFavouritesPrompt}
            iconProps={{ iconName: "Filter" }}
            onChange={_onFilterChanged}
          />
          <div>{status}</div>
          <FocusZone direction={FocusZoneDirection.vertical}>
            {myFavouriteItems.length > 0 ? (
              <List items={myFavouriteItems} onRenderCell={_onRenderCell} />
            ) : (
              <MessageBar
                messageBarType={MessageBarType.warning}
                isMultiline={false}
              >
                {strings.NoFavouritesLabel}
              </MessageBar>
            )}
          </FocusZone>
        </div>
      </Panel>
      <Dialog
        hidden={!showDialog}
        onDismiss={_hideDialog}
        dialogContentProps={{
          type: DialogType.largeHeader,
          title: dialogTitle,
        }}
        modalProps={{
          titleAriaId: "myFavDialog",
          subtitleAriaId: "myFavDialog",
          isBlocking: false,
          containerClassName: "ms-dialogMainOverride",
        }}
      >
        <div>{status}</div>
        <TextField
          label={strings.TitleFieldName}
          onChange={_setItemInContext}
          value={itemInContext.Title}
          name="Title"
        />
        <TextField
          label={strings.DescriptionFieldName}
          multiline
          rows={4}
          onChange={_setItemInContext}
          value={itemInContext.Description}
          name="Description"
        />
        <DialogFooter>
          <PrimaryButton
            onClick={(_event) => _saveMyFavourite()}
            disabled={disableButtons}
            text={strings.SaveButtonLabel}
          />
          <DefaultButton
            onClick={_hideDialog}
            disabled={disableButtons}
            text={strings.CancelButtonLabel}
          />
        </DialogFooter>
      </Dialog>
    </div>
  );
};

export default MyFavouritesTopBar;
