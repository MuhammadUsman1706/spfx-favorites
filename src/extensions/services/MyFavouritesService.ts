import { IMyFavouritesService } from "./IMyFavouritesService";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { IMyFavouritesTopBarProps } from "../myFavourites/components/MyFavouritesTopBar/IMyFavouritesTopBarProps";
import { SPFI, spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";

import { IItemAddResult, IItemUpdateResult } from "@pnp/sp/items";
import { IWebEnsureUserResult } from "@pnp/sp/site-users/";
import { IMyFavouriteItem } from "../interfaces/IMyFavouriteItem";
import { Log } from "@microsoft/sp-core-library";

const LOG_SOURCE: string = "CC_MyFavourites_ApplicationCustomizer";
const FAVOURITES_LIST_NAME: string = "Favorites";

export class MyFavouritesService implements IMyFavouritesService {
  private _context: ApplicationCustomizerContext;
  private _currentWebUrl: string;
  private _sessionStorageKey: string = "MyFavourites_";
  private sp: SPFI;
  constructor(_props: IMyFavouritesTopBarProps) {
    console.log("Favorites extension has loaded successfully!");
    this._context = _props.context;
    this._currentWebUrl = this._context.pageContext.web.absoluteUrl;
    this._sessionStorageKey += this._currentWebUrl;
    this.sp = spfi().using(SPFx(this._context));
  }

  public async getMyFavourites(
    tryFromCache: boolean
  ): Promise<IMyFavouriteItem[]> {
    let myFavourites: IMyFavouriteItem[] = [];
    if (tryFromCache) {
      myFavourites = this._fetchFromSessionStorge();
      if (myFavourites.length) {
        return myFavourites;
      }
    }
    myFavourites = await this._fetchFromSPList();
    let favInCache: string = JSON.stringify(myFavourites);
    window.sessionStorage.setItem(this._sessionStorageKey, favInCache);
    return myFavourites;
  }

  public async saveFavourite(
    favouriteItem: IMyFavouriteItem
  ): Promise<boolean> {
    return this.sp.web.lists
      .getByTitle(FAVOURITES_LIST_NAME)
      .items.add({
        Title: favouriteItem.Title,
        Description: favouriteItem.Description,
        ItemUrl: window.location.href,
      })
      .then(
        async (result: IItemAddResult): Promise<boolean> => {
          // let addedItem: IMyFavouriteItem = result.data;
          await this.getMyFavourites(false);
          return true;
        },
        (error: any): boolean => {
          return false;
        }
      );
  }

  public async deleteFavourite(favouriteItemId: number): Promise<boolean> {
    return this.sp.web.lists
      .getByTitle(FAVOURITES_LIST_NAME)
      .items.getById(favouriteItemId)
      .delete()
      .then(
        async (): Promise<boolean> => {
          await this.getMyFavourites(false);
          return true;
        },
        (_error: any): boolean => {
          return false;
        }
      );
  }

  public async updateFavourite(
    favouriteItem: IMyFavouriteItem
  ): Promise<boolean> {
    return this.sp.web.lists
      .getByTitle(FAVOURITES_LIST_NAME)
      .items.getById(Number(favouriteItem.Id))
      .update({
        Title: favouriteItem.Title,
        Description: favouriteItem.Description,
      })
      .then(
        async (_result: IItemUpdateResult): Promise<boolean> => {
          await this.getMyFavourites(false);
          return true;
        },
        (_error: any): boolean => {
          return false;
        }
      );
  }

  private _fetchFromSessionStorge(): IMyFavouriteItem[] {
    let result: IMyFavouriteItem[] = [];
    let stringResult: string =
      window.sessionStorage.getItem(this._sessionStorageKey) || "";
    if (stringResult) {
      try {
        Log.info(LOG_SOURCE, "Fetched favourites from cache");
        result = JSON.parse(stringResult);
      } catch (error) {
        Log.error(LOG_SOURCE, error);
      }
    }
    return result;
  }

  private async _fetchFromSPList(): Promise<IMyFavouriteItem[]> {
    const currentUserId: number = await this._getUserId();
    return (
      this.sp.web.lists
        .getByTitle(FAVOURITES_LIST_NAME)
        .items.select("Id", "Title", "ItemUrl", "Description")
        .filter("Author eq " + currentUserId)()
        // .get()
        .then((myFavourites: IMyFavouriteItem[]) => {
          Log.info(LOG_SOURCE, "Fetched favourites from list");
          return myFavourites;
        })
        .catch((error: any) => {
          Log.error(LOG_SOURCE, error);
          return [];
        })
    );
  }

  private _getUserId(): Promise<number> {
    return this.sp.web
      .ensureUser(this._context.pageContext.user.email)
      .then((result: IWebEnsureUserResult) => {
        return result.data.Id;
      });
  }
}
