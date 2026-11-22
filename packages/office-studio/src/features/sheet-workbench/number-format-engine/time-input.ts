import type { ParseOptions } from './parse-value';
import type { LocaleData, ParseData } from './types';
import { defaultLocale, getLocale } from './locale';
import { normalizeInputDigits, parseNumber } from './number-input';

const SECONDS_PER_DAY = 86400;
const MAX_SECOND_DECIMALS = 10;

export interface DurationParseOptions extends ParseOptions {
    unit?: 'hours' | 'minutes' | 'seconds';
}

interface SecondFraction {
    text: string;
    decimals: number;
}

function normalizeTimeInput(value: string, locale: LocaleData): string {
    const source = normalizeInputDigits(value).trim().replace(/：/g, ':');
    return locale.decimal === '.' ? source : source.replace(locale.decimal, '.');
}

function normalizeAmPm(value: string): string {
    return value.replace(/[\s.]/g, '').toLowerCase();
}

function secondFraction(value: string | undefined): SecondFraction | null {
    if (!value) {
        return { text: '', decimals: 0 };
    }
    const digits = value.slice(1);
    if (!/^\d+$/.test(digits) || digits.replace(/0+$/, '').length > MAX_SECOND_DECIMALS) {
        return null;
    }
    return { text: value, decimals: Math.min(MAX_SECOND_DECIMALS, digits.length) };
}

function fractionPattern(fraction: SecondFraction): string {
    return fraction.decimals ? `.${'0'.repeat(fraction.decimals)}` : '';
}

function secondsToSerial(wholeSeconds: number, fraction: SecondFraction, negative = false): number | null {
    if (!Number.isSafeInteger(wholeSeconds) || wholeSeconds < 0) {
        return null;
    }
    // Validate the combined seconds, since adding a tiny fraction to a large clock can lose it before division.
    const parsed = parseNumber(`${negative ? '-' : ''}${wholeSeconds}${fraction.text}`, { locale: 'en' });
    if (!parsed) {
        return null;
    }
    const serial = parsed.v / SECONDS_PER_DAY;
    return Number.isFinite(serial) && (serial !== 0 || parsed.v === 0) ? serial : null;
}

function readAmPmSuffix(value: string, locale: LocaleData): 'a' | 'p' | '' | null {
    const suffix = normalizeAmPm(value);
    if (!suffix) {
        return '';
    }
    if (suffix === normalizeAmPm(locale.ampm[0]) || suffix === 'a' || suffix === 'am') {
        return 'a';
    }
    if (suffix === normalizeAmPm(locale.ampm[1]) || suffix === 'p' || suffix === 'pm') {
        return 'p';
    }
    return null;
}

/** Parse a localized clock, including a complete next-day midnight and prefix/suffix AM/PM. */
export function parseTime(value: string, options: ParseOptions = {}): ParseData<number> | null {
    const locale = getLocale(options.locale || '') ?? defaultLocale;
    let source = normalizeTimeInput(value, locale);
    let prefixAmPm: 'a' | 'p' | '' = '';
    const prefixes: Array<[string, 'a' | 'p']> = [
        [locale.ampm[0], 'a'], [locale.ampm[1], 'p'],
        ['a.m.', 'a'], ['p.m.', 'p'], ['AM', 'a'], ['PM', 'p'],
    ];
    prefixes.sort((left, right) => right[0].length - left[0].length);
    for (const [prefix, type] of prefixes) {
        if (!prefix || !source.toLowerCase().startsWith(prefix.toLowerCase())) {
            continue;
        }
        const remaining = source.slice(prefix.length);
        if (!/^\s*\d/.test(remaining) || /[a-z.]$/i.test(prefix) && !/^\s/.test(remaining)) {
            continue;
        }
        prefixAmPm = type;
        source = remaining.trimStart();
        break;
    }

    const match = /^(\d{1,2})(?::(\d{1,2}))?(?::(\d{1,2}))?(\.\d+)?(:)?(.*)$/.exec(source);
    if (!match) {
        return null;
    }
    const hoursText = match[1];
    if (hoursText === undefined) {
        return null;
    }
    let minutesText = match[2];
    let secondsText = match[3];
    const fraction = secondFraction(match[4]);
    const emptyFinalSegment = !!match[5];
    const suffixText = match[6] ?? '';
    const suffixAmPm = readAmPmSuffix(suffixText, locale);
    if (!fraction || suffixAmPm === null || prefixAmPm && suffixAmPm) {
        return null;
    }
    if (emptyFinalSegment) {
        if (secondsText !== undefined || fraction.text) {
            return null;
        }
        minutesText ??= '0';
        secondsText = '0';
    }
    const ampm = prefixAmPm || suffixAmPm;
    if (fraction.text && secondsText === undefined || !ampm && minutesText === undefined) {
        return null;
    }
    // An unseparated letter after a bare hour is text; a localized prefix may be adjacent to its clock.
    if (suffixAmPm && minutesText === undefined && !/^\s/.test(suffixText)) {
        return null;
    }
    let hours = Number(hoursText);
    const minutes = Number(minutesText ?? 0);
    const seconds = Number(secondsText ?? 0);
    if (hours > 24 || minutes > 59 || seconds > 59) {
        return null;
    }
    if (ampm) {
        if (hours < 1 || hours > 12) {
            return null;
        }
        hours = hours % 12 + (ampm === 'p' ? 12 : 0);
    } else if (hours === 24 && (minutes || seconds || Number(fraction.text))) {
        return null;
    }
    const serial = secondsToSerial(hours * 3600 + minutes * 60 + seconds, fraction);
    if (serial === null || serial > 1 || serial === 1 && hours !== 24) {
        return null;
    }
    const secondsFormat = secondsText === undefined ? '' : `:ss${fractionPattern(fraction)}`;
    const clock = `${hoursText.length === 2 ? 'hh' : 'h'}:mm${secondsFormat}`;
    return {
        v: serial,
        z: prefixAmPm ? `AM/PM ${clock}` : `${clock}${ampm ? ' AM/PM' : ''}`,
    };
}

/** Parse elapsed hours, minutes, or seconds without folding the elapsed part into a 24-hour clock. */
export function parseDuration(value: string, options: DurationParseOptions = {}): ParseData<number> | null {
    const locale = getLocale(options.locale || '') ?? defaultLocale;
    let source = normalizeTimeInput(value, locale);
    let negative = false;
    const signs: Array<[string, boolean]> = [
        [locale.negative, true], ['−', true], ['－', true], ['-', true],
        [locale.positive, false], ['＋', false], ['+', false],
    ];
    for (const [sign, signNegative] of signs) {
        if (sign && source.startsWith(sign)) {
            negative = signNegative;
            source = source.slice(sign.length);
            break;
        }
    }
    const unit = options.unit ?? 'hours';
    if (options.unit === undefined && !source.includes(':')) {
        return null;
    }
    const pattern = unit === 'seconds'
        ? /^(\d*)(\.\d+)?$/
        : unit === 'minutes'
            ? /^(\d+)(?::(\d{1,2}))?(\.\d+)?$/
            : /^(\d+)(?::(\d{1,2}))?(?::(\d{1,2}))?(\.\d+)?$/;
    const match = pattern.exec(source);
    if (!match || !match[1] && !match[2]) {
        return null;
    }
    const elapsed = Number(match[1] || 0);
    if (!Number.isSafeInteger(elapsed)) {
        return null;
    }
    let wholeSeconds: number;
    let format: string;
    let fractionText: string | undefined;
    if (unit === 'seconds') {
        wholeSeconds = elapsed;
        fractionText = match[2];
        format = '[s]';
    } else if (unit === 'minutes') {
        const seconds = Number(match[2] ?? 0);
        if (seconds > 59 || match[3] && match[2] === undefined) {
            return null;
        }
        wholeSeconds = elapsed * 60 + seconds;
        fractionText = match[3];
        format = match[2] === undefined ? '[m]' : '[m]:ss';
    } else {
        const minutes = Number(match[2] ?? 0);
        const seconds = Number(match[3] ?? 0);
        if (minutes > 59 || seconds > 59 || match[4] && match[3] === undefined) {
            return null;
        }
        wholeSeconds = elapsed * 3600 + minutes * 60 + seconds;
        fractionText = match[4];
        format = match[2] === undefined ? '[h]' : match[3] === undefined ? '[h]:mm' : '[h]:mm:ss';
    }
    const fraction = secondFraction(fractionText);
    if (!fraction) {
        return null;
    }
    const serial = secondsToSerial(wholeSeconds, fraction, negative);
    return serial === null ? null : { v: serial, z: format + fractionPattern(fraction) };
}
