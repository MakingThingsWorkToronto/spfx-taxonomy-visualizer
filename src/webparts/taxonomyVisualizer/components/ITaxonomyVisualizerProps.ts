import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { TaxonomyService } from '../../../services/TaxonomyService';
import { DisplayMode } from '@microsoft/sp-core-library';
import { IColumnBreakpoints } from '../../../models/IColumnBreakpoints';

export interface ITaxonomyVisualizerProps {
  title: string;
  displayMode: DisplayMode;
  updateProperty: (value: string) => void;
  theme: IReadonlyTheme;
  service: TaxonomyService;
  termSetId:string;
  linkTemplate:string;
  breakpoints:IColumnBreakpoints[];
  levels:number;
  lcid:number;
  domElement: HTMLElement;
}
