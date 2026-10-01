declare interface IGeekLauncherCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'GeekLauncherCommandSetStrings' {
  const strings: IGeekLauncherCommandSetStrings;
  export = strings;
}
