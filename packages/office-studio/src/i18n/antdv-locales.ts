import type { ConfigProviderProps } from 'antdv-next'
import arEG from 'antdv-next/locale/ar_EG'
import caES from 'antdv-next/locale/ca_ES'
import deDE from 'antdv-next/locale/de_DE'
import enUS from 'antdv-next/locale/en_US'
import esES from 'antdv-next/locale/es_ES'
import faIR from 'antdv-next/locale/fa_IR'
import frFR from 'antdv-next/locale/fr_FR'
import idID from 'antdv-next/locale/id_ID'
import itIT from 'antdv-next/locale/it_IT'
import jaJP from 'antdv-next/locale/ja_JP'
import koKR from 'antdv-next/locale/ko_KR'
import plPL from 'antdv-next/locale/pl_PL'
import ptBR from 'antdv-next/locale/pt_BR'
import ruRU from 'antdv-next/locale/ru_RU'
import skSK from 'antdv-next/locale/sk_SK'
import viVN from 'antdv-next/locale/vi_VN'
import zhCN from 'antdv-next/locale/zh_CN'
import zhHK from 'antdv-next/locale/zh_HK'
import zhTW from 'antdv-next/locale/zh_TW'
import 'dayjs/locale/ar'
import 'dayjs/locale/ca'
import 'dayjs/locale/de'
import 'dayjs/locale/es'
import 'dayjs/locale/fa'
import 'dayjs/locale/fr'
import 'dayjs/locale/id'
import 'dayjs/locale/it'
import 'dayjs/locale/ja'
import 'dayjs/locale/ko'
import 'dayjs/locale/pl'
import 'dayjs/locale/pt-br'
import 'dayjs/locale/ru'
import 'dayjs/locale/sk'
import 'dayjs/locale/vi'
import 'dayjs/locale/zh-cn'
import 'dayjs/locale/zh-hk'
import 'dayjs/locale/zh-tw'
import type { OfficeLocale } from './regional-settings.ts'

type AntdvLocale = NonNullable<ConfigProviderProps['locale']>
const locales = {
  'ar-SA': arEG, 'ca-ES': caES, 'de-DE': deDE, 'en-US': enUS, 'es-ES': esES, 'fa-IR': faIR,
  'fr-FR': frFR, 'id-ID': idID, 'it-IT': itIT, 'ja-JP': jaJP, 'ko-KR': koKR, 'pl-PL': plPL,
  'pt-BR': ptBR, 'ru-RU': ruRU, 'sk-SK': skSK, 'vi-VN': viVN, 'zh-CN': zhCN, 'zh-HK': zhHK, 'zh-TW': zhTW,
} satisfies Record<OfficeLocale, AntdvLocale>

export function getAntdvLocale(locale: OfficeLocale): AntdvLocale { return locales[locale] }
