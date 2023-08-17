declare interface IScAdminWebPartStrings {
  FractionDigitsFieldLabel: string;
  FractionDigitsFieldValue: number;
  SearchFileTypesFieldDescription: string;
  SearchFileTypesFieldLabel: string;
  SearchFileTypesFieldValue: string;
  SearchMonthsFieldLabel: string;
  SearchMonthsFieldValue: number;
  SearchTermsFieldDescription: string;
  SearchTermsFieldLabel: string;
  SearchTermsFieldValue: string;
  TimeFormatFieldDescription: string;
  TimeFormatFieldLabel: string;
  TimeFormatFieldValue: string;
}

declare module 'ScAdminWebPartStrings' {
  const strings: IScAdminWebPartStrings;
  export = strings;
}
