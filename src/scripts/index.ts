export * from "./documentRetention";
export * from "./exportCSV";
export * from "./lists";
export * from "./securityGroups";
export * from "./sites";

export interface IScript {
    init: any;
    name: string;
    description: string;
}