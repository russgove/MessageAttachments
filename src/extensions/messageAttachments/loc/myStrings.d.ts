declare interface IMessageAttachmentsCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'MessageAttachmentsCommandSetStrings' {
  const strings: IMessageAttachmentsCommandSetStrings;
  export = strings;
}
