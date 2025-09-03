declare interface ISendForSignatureCommandCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'SendForSignatureCommandCommandSetStrings' {
  const strings: ISendForSignatureCommandCommandSetStrings;
  export = strings;
}
