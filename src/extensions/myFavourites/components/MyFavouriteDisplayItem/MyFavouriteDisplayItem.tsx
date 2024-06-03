import * as React from "react";
import { IMyFavouriteDisplayItemProps } from "./IMyFavouriteDisplayItemProps";
import { PrimaryButton } from "office-ui-fabric-react/lib/Button";
import { Link } from "office-ui-fabric-react/lib/Link";
import { Spinner, SpinnerSize } from "office-ui-fabric-react/lib/Spinner";
import { css } from "@uifabric/utilities/lib/css";

import styles from "../MyFavourites.module.scss";
import * as strings from "MyFavouritesApplicationCustomizerStrings";

const MyFavouriteDisplayItem: React.FC<IMyFavouriteDisplayItemProps> = (
  props
) => {
  const [status, setStatus] = React.useState<JSX.Element>(<span></span>);
  const [disableButtons, setDisableButtons] = React.useState<boolean>(false);

  const firstLetter: string | undefined =
    props.displayItem.Title?.charAt(0).toUpperCase();

  const deleteFavourite = async (): Promise<void> => {
    setStatus(<Spinner size={SpinnerSize.small} />);
    setDisableButtons(true);
    await props.deleteFavourite(Number(props.displayItem.Id));
    setStatus(<span></span>);
    setDisableButtons(false);
  };

  const editFavourite = () => {
    setStatus(<Spinner size={SpinnerSize.small} />);
    setDisableButtons(true);
    props.editFavoutite(props.displayItem);
    setStatus(<span></span>);
    setDisableButtons(false);
  };

  return (
    <div className={`${styles.ccitemContent}`}>
      <Link href={props.displayItem.ItemUrl} className={styles.ccRow}>
        <div className={css("ms-font-su", styles.ccInitials)}>
          {firstLetter}
        </div>
        <div className={styles.ccitemName}>
          <span className={"ms-font-l"}>{props.displayItem.Title}</span>
        </div>
        <div className={styles.ccitemDesc}>{props.displayItem.Description}</div>
      </Link>
      <div className={styles.ccitemDesc}>
        <PrimaryButton
          data-automation-id="btnEdit"
          iconProps={{ iconName: "Edit" }}
          text={strings.EditButtonLabel}
          disabled={disableButtons}
          onClick={editFavourite}
          className={styles.ccButton}
        />
        <PrimaryButton
          data-automation-id="btnDel"
          iconProps={{ iconName: "ErrorBadge" }}
          text={strings.DeleteButtonLabel}
          disabled={disableButtons}
          onClick={deleteFavourite}
          className={styles.ccButton}
        />
        <div className={styles.ccStatus}>{status}</div>
      </div>
    </div>
  );
};

export default MyFavouriteDisplayItem;
