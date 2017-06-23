declare interface IHeaderFooterStrings {
  Title: string;
}

declare module 'headerFooterStrings' {
  const strings: IHeaderFooterStrings;
  export = strings;
}
