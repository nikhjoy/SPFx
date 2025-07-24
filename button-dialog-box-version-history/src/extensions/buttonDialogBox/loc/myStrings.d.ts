declare interface IButtonDialogBoxCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'ButtonDialogBoxCommandSetStrings' {
  const strings: IButtonDialogBoxCommandSetStrings;
  export = strings;
}
