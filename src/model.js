import {Workbook} from './model-base.js';
import {installPivotModel} from './pivots/model.js';
installPivotModel(Workbook);
export * from './model-base.js';
export {PivotReport,PivotTableCollection} from './pivots/model.js';
