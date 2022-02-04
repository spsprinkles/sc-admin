declare interface IAdminMenuCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'AdminMenuCommandSetStrings' {
  const strings: IAdminMenuCommandSetStrings;
  export = strings;
}
