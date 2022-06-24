declare interface IHelloSignCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'HelloSignCommandSetStrings' {
  const strings: IHelloSignCommandSetStrings;
  export = strings;
}
