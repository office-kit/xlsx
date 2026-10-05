// UI language. English strings are the keys' source of truth (en.ts); ja.ts
// translates them. Missing translations fall back to English so a new string
// never renders as a raw key.

import { en, type MessageKey } from './en.ts';
import { ja } from './ja.ts';

export type Locale = 'en' | 'ja';

const dictionaries: Record<Locale, Partial<Record<MessageKey, string>>> = { en, ja };

function detect(): Locale {
  if (typeof navigator === 'undefined') return 'en';
  return navigator.language.toLowerCase().startsWith('ja') ? 'ja' : 'en';
}

class I18n {
  locale = $state<Locale>(detect());
}

export const i18n = new I18n();

/** Translate `key`, substituting `{name}` placeholders from `params`. */
export function t(key: MessageKey, params?: Record<string, string | number>): string {
  const text = dictionaries[i18n.locale][key] ?? en[key];
  if (!params) return text;
  return text.replace(/\{(\w+)\}/g, (_: string, name: string) => String(params[name] ?? `{${name}}`));
}

export type { MessageKey };
