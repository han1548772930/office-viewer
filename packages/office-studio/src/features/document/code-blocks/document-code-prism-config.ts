/** Configure the lexer before its module installs browser or Worker hooks. */
export const DOCUMENT_CODE_PRISM_CONFIG = { manual: true, disableWorkerMessageHandler: true } as const

const scope = globalThis as typeof globalThis & { Prism?: { manual?: boolean; disableWorkerMessageHandler?: boolean } }
if (scope.Prism) Object.assign(scope.Prism, DOCUMENT_CODE_PRISM_CONFIG)
else scope.Prism = { ...DOCUMENT_CODE_PRISM_CONFIG }
