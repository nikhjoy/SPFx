declare interface IHiCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'HiCommandSetStrings' {
  const strings: IHiCommandSetStrings;
  export = strings;
}
