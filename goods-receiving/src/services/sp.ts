import { spfi, SPFI, SPFx } from '@pnp/sp';
import { ISPFXContext } from '@pnp/sp/behaviors/spfx';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/files';
import '@pnp/sp/folders';
import '@pnp/sp/site-users/web';
import '@pnp/sp/site-groups';
import '@pnp/sp/security/list';
import '@pnp/sp/profiles';

/**
 * Creates the PnPjs entry point. Authentication uses the signed-in Microsoft 365
 * user through the SPFx context; no extra API permissions are required.
 */
export function createSp(context: ISPFXContext, siteUrl: string): SPFI {
  return spfi(siteUrl).using(SPFx(context));
}
