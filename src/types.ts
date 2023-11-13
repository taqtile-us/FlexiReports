export interface IMaster {
  entityName: string,
  fieldName: string,
  address: string,
  type: string
  alignment: IAlignment
}

export interface IFormula {
  formula: string;
  address: string;
  alignment: IAlignment
}

export interface IFormulas {
  rowFormulas: IFormula[];
  columnFormulas: IFormula[];
}

interface IAlignment {
  horizontal: string, vertical: string
}

interface IStaticVariable {
  value: string;
  address: string;
  alignment: IAlignment
}

export interface IStaticVariables {
  [key: string]: IStaticVariable;
}


export interface IDetail {
  entityName: string;
  fieldName: string;
  address: string;
  type: string;
  alignment: IAlignment
}

export interface IDetails {
  [key: string]: IDetail[];
}
