// Threaded-comment authors (`xl/persons/person.xml`).

import { newOfficeGuid } from '../worksheet/threaded-comments.js';

export interface Person {
  /** Name shown on the comment card. */
  displayName: string;
  /** GUID in braces; threaded comments reference their author by this id. */
  id: string;
  /** Account identifier at the identity provider (e-mail, SID …). */
  userId?: string;
  /** Identity provider: `AD`, `Windows Live`, or `None` for a local user. */
  providerId?: string;
}

/**
 * Build a person with a new id. Without an account Excel records the display
 * name as the user id under provider `None`, so that is the default here too.
 */
export function makePerson(opts: { displayName: string; userId?: string; providerId?: string }): Person {
  return {
    displayName: opts.displayName,
    id: newOfficeGuid(),
    userId: opts.userId ?? opts.displayName,
    providerId: opts.providerId ?? 'None',
  };
}
