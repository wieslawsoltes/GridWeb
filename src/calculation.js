import {CalculationEngine as BaseCalculationEngine} from './calculation-base.js';
import {parseFormula} from './parser.js';
import {error,isError,number,text,truth,scalar,matrix} from './errors.js';
import {parseRange,cellAddress,MAX_ROWS,MAX_COLUMNS,MAX_OPERATION_CELLS} from './address.js';
import {criteriaPredicate} from './functions.js';
import {databaseAggregate,DATABASE_FUNCTIONS} from './functions-database.js';
import {referenceAggregate} from './calculation-aggregate.js';
import {OMITTED,LAMBDA_HELPERS,evaluateLambdaHelper} from './calculation-lambda.js';
import {shiftFormula} from './address.js';
import {r1c1ToA1} from './references.js';
export {isFormula,literalValue} from './calculation-base.js';
const aggregate=new Set('SUM AVERAGE MIN MAX COUNT COUNTA PRODUCT SUMSQ MEDIAN STDEV STDEVP STDEV.S STDEV.P VAR VARP VAR.S VAR.P AVEDEV DEVSQ GEOMEAN HARMEAN MODE MODE.SNGL MODE.MULT SKEW SKEW.P KURT'.split(' '));
const numericAggregate=new Set('AVEDEV DEVSQ GEOMEAN HARMEAN MODE MODE.SNGL MODE.MULT SKEW SKEW.P KURT'.split(' '));
const scalarFunctions=new Set('ABS ACOS ACOSH ASIN ASINH ATAN ATAN2 ATANH COS COSH SIN SINH TAN TANH DEGREES RADIANS EXP LN LOG LOG10 SQRT SIGN INT TRUNC ROUND ROUNDUP ROUNDDOWN MROUND CEILING.MATH FLOOR.MATH MOD QUOTIENT POWER EVEN ODD FACT LEN LEFT RIGHT MID LOWER UPPER PROPER TRIM CLEAN CHAR CODE UNICHAR UNICODE REPLACE SUBSTITUTE FIND SEARCH EXACT VALUE NUMBERVALUE YEAR MONTH DAY HOUR MINUTE SECOND DATE TIME DAYS EDATE EOMONTH ISOWEEKNUM WEEKNUM STANDARDIZE GAMMA GAMMALN GAMMALN.PRECISE NORM.DIST NORM.S.DIST NORM.INV NORM.S.INV GAMMA.DIST GAMMA.INV BETA.DIST BETA.INV T.DIST T.DIST.RT T.DIST.2T T.INV T.INV.2T F.DIST F.DIST.RT F.INV F.INV.RT CHISQ.DIST CHISQ.DIST.RT CHISQ.INV CHISQ.INV.RT BINOM.DIST BINOM.INV POISSON.DIST EXPON.DIST WEIBULL.DIST ERF ERFC ERF.PRECISE ERFC.PRECISE BITAND BITOR BITXOR BITLSHIFT BITRSHIFT BASE DECIMAL DELTA GESTEP'.split(' '));
function lifted(args,fn){
  if(!args.some(Array.isArray))return fn(...args);
  const arrays=args.map(matrix),h=Math.max(...arrays.map(a=>a.length)),w=Math.max(...arrays.map(a=>a[0]?.length??0));
  if(!h||!w||h*w>MAX_OPERATION_CELLS||arrays.some(a=>a.length!==1&&a.length!==h||a.some(r=>r.length!==1&&r.length!==w)))return error('#VALUE!','Array dimensions differ');
  return Array.from({length:h},(_,r)=>Array.from({length:w},(_,c)=>{try{return fn(...arrays.map(a=>a[a.length===1?0:r][a[0].length===1?0:c]));}catch(e){return isError(e)?e:error('#VALUE!',e.message);}}));
}
export class CalculationEngine extends BaseCalculationEngine {
  get FunctionNames(){return [...new Set([...super.FunctionNames,'AGGREGATE','MAKEARRAY','ISOMITTED'])].sort();}
  Evaluate(formula,context){
    const root=!this._lambdaBudget;if(root)this._lambdaBudget={calls:0,depth:0};
    try{return super.Evaluate(formula,context);}finally{if(root)this._lambdaBudget=null;}
  }
  _invoke(lambda,values){
    if(lambda?.type!=='lambda'||values.length!==lambda.parameters.length)return error('#VALUE!','Incorrect Parameters');
    const budget=this._lambdaBudget??{calls:0,depth:0};
    if(++budget.calls>MAX_OPERATION_CELLS*4||budget.depth>=128)return error('#NUM!','Lambda evaluation limit');
    const vars=new Map(lambda.context.vars),omitted=new Set(lambda.context.omitted??[]);
    lambda.parameters.forEach((p,i)=>{vars.set(p,values[i]===OMITTED?null:values[i]);omitted.delete(p);if(values[i]===OMITTED)omitted.add(p);});
    budget.depth++;
    try{return this._eval(lambda.body,{...lambda.context,vars,omitted,names:new Set(),depth:lambda.context.depth+1});}finally{budget.depth--;}
  }
  /** Resolve reference-valued expressions without collapsing their origin/type. */
  _reference(node,ctx,depth=0){
    if(!node||depth>64)return null;
    if(node.type==='ref')return node;
    if(node.type==='name'&&!ctx.vars?.has(node.name)){
      const value=this.Workbook._names.get(node.name);
      if(typeof value==='string'&&value.startsWith('='))return this._reference(parseFormula(value),ctx,depth+1);
    }
    if(node.type==='call'&&node.name==='INDIRECT'){
      const a=node.args,source=text(this._eval(a[0],ctx)),useA1=a[1]==null||a[1].value===null||truth(this._eval(a[1],ctx));
      try{return {...parseRange(useA1?source:r1c1ToA1(source,{row:ctx.row,column:ctx.col})),type:'ref'};}catch{throw error('#REF!');}
    }
    if(node.type==='call'&&node.name==='OFFSET'){
      const a=node.args,base=this._reference(a[0],ctx,depth+1);if(!base)throw error('#VALUE!');const ev=i=>Math.trunc(number(this._eval(a[i],ctx)));
      const r=ev(1),c=ev(2),h=a[3]==null||a[3].value===null?base.r2-base.r1+1:ev(3),w=a[4]==null||a[4].value===null?base.c2-base.c1+1:ev(4);
      const ref={...base,r1:base.r1+r,c1:base.c1+c,r2:base.r1+r+h-1,c2:base.c1+c+w-1};
      if(h<1||w<1||ref.r1<0||ref.c1<0||ref.r2>=MAX_ROWS||ref.c2>=MAX_COLUMNS)throw error('#REF!');return ref;
    }
    return null;
  }
  _evaluate(node,ctx){
    if(!node||ctx.depth>128)return error('#NUM!','Evaluation nesting limit');
    const next={...ctx,depth:ctx.depth+1},ev=n=>this._eval(n,next);
    if(node.type==='unary'&&node.op==='@'){
      const ref=this._reference(node.value,next);
      if(ref){const s=this._sheet(ref.sheet,ctx.sheet);if(!s)return error('#REF!');
        const r=ref.r1===ref.r2?ref.r1:ctx.row,c=ref.c1===ref.c2?ref.c1:ctx.col;
        if(r<ref.r1||r>ref.r2||c<ref.c1||c>ref.c2||ref.r1!==ref.r2&&ref.c1!==ref.c2)return error('#VALUE!','No unique implicit intersection');
        return this.GetValue(s,r,c,next);
      }return scalar(ev(node.value));
    }
    if(node.type==='invoke')return this._invoke(ev(node.callee),node.args.map(a=>a.omitted?OMITTED:ev(a)));
    if(node.type!=='call')return super._evaluate(node,ctx);
    const {name,args}=node;
    if(name==='LET'){
      if(args.length<3||args.length%2===0||args.length>253)return error('#VALUE!','LET argument count');
      const vars=new Map(ctx.vars),omitted=new Set(ctx.omitted??[]);
      for(let i=0;i<args.length-1;i+=2){if(args[i].type!=='name')return error('#NAME?','Invalid LET binding');const value=this._eval(args[i+1],{...next,vars,omitted});vars.set(args[i].name,value);omitted.delete(args[i].name);}
      return this._eval(args.at(-1),{...next,vars,omitted});
    }
    if(name==='LAMBDA'){
      const parameters=args.slice(0,-1);
      if(!args.length||parameters.length>253||parameters.some(p=>p.type!=='name'||p.name.includes('.'))||new Set(parameters.map(p=>p.name)).size!==parameters.length)return error('#VALUE!','Invalid LAMBDA parameters');
      return {type:'lambda',parameters:parameters.map(p=>p.name),body:args.at(-1),context:next};
    }
    if(name==='ISOMITTED')return args.length===1?(args[0].type==='name'&&!!ctx.omitted?.has(args[0].name)):error('#VALUE!','ISOMITTED argument count');
    if(LAMBDA_HELPERS.has(name))return evaluateLambdaHelper(this,name,args,ev);
    if(name==='AGGREGATE'||name==='SUBTOTAL')return referenceAggregate(this,name,args,next,ev);
    if(Object.hasOwn(DATABASE_FUNCTIONS,name)){
      if(args.length!==3)return error('#VALUE!','Database function argument count');
      const data=ev(args[0]),field=ev(args[1]),criteria=ev(args[2]),ref=this._reference(args[2],next);
      const criteriaSheet=ref?this._sheet(ref.sheet,ctx.sheet):null;
      const computed=criteriaSheet?(dataRow,criteriaRow,column)=>{
        const row=ref.r1+criteriaRow,col=ref.c1+column,record=criteriaSheet._cells.get(row*MAX_COLUMNS+col);
        if(record?.literal||typeof record?.input!=='string'||!record.input.startsWith('='))throw error('#VALUE!','Calculated criteria require a formula');
        const value=this.Evaluate(shiftFormula(record.input,dataRow-1,0),{...next,sheet:criteriaSheet,row,col});
        return truth(value);
      }:null;
      return databaseAggregate(this.Functions,criteriaPredicate,name,data,field,criteria,computed);
    }
    if(name==='ADDRESS'){
      if(args.length<2||args.length>5)return error('#VALUE!','ADDRESS argument count');
      const r=Math.trunc(number(ev(args[0]))),c=Math.trunc(number(ev(args[1]))),mode=args[2]==null||args[2].value===null?1:Math.trunc(number(ev(args[2]))),a1=args[3]==null||args[3].value===null||truth(ev(args[3]));
      if(mode<1||mode>4||r<1||c<1||r>MAX_ROWS||c>MAX_COLUMNS)return error('#VALUE!');
      const ar=mode<=2,ac=mode===1||mode===3;
      const address=a1?cellAddress(r-1,c-1).replace(/^([A-Z]+)(\d+)$/,(_,col,row)=>(ac?'$':'')+col+(ar?'$':'')+row):'R'+(ar?r:`[${r}]`)+'C'+(ac?c:`[${c}]`);
      const sheet=args[4]&&args[4].value!==null?text(ev(args[4])):null;return sheet==null?address:(/^[A-Za-z_][\w.]*$/.test(sheet)?sheet:"'"+sheet.replace(/'/g,"''")+"'")+'!'+address;
    }
    if(name==='INDIRECT'||name==='OFFSET'){
      if(args.length<(name==='INDIRECT'?1:3)||args.length>(name==='INDIRECT'?2:5))return error('#VALUE!');
      return this._read(this._reference(node,next),next);
    }
    if(name==='ROW'||name==='COLUMN'){
      if(args.length>1)return error('#VALUE!');if(!args.length)return(name==='ROW'?ctx.row:ctx.col)+1;
      const ref=this._reference(args[0],next);if(!ref)return error('#VALUE!');if(!this._sheet(ref.sheet,ctx.sheet))return error('#REF!');
      const row=name==='ROW',start=row?ref.r1:ref.c1,count=(row?ref.r2:ref.c2)-start+1;
      if(count>MAX_OPERATION_CELLS)return error('#NUM!');
      return count===1?start+1:row?Array.from({length:count},(_,i)=>[start+i+1]):[Array.from({length:count},(_,i)=>start+i+1)];
    }
    if(name==='IFERROR'||name==='IFNA'){
      if(args.length!==2)return error('#VALUE!');const v=ev(args[0]),matches=x=>isError(x)&&(name==='IFERROR'||x.code==='#N/A');
      if(!matrix(v).some(r=>r.some(matches)))return v;
      return lifted([v,ev(args[1])],(a,b)=>matches(a)?b:a);
    }
    const fn=this.Functions.get(name);
    if(fn&&aggregate.has(name)){
      if(!args.length||args.length>255)return error('#VALUE!','Aggregate argument count');
      return fn(...args.map(arg=>{const ref=this._reference(arg,next);if(ref){const value=this._read(ref,next,true);return Array.isArray(value)?value:[[value]];}const v=ev(arg);if(arg.type==='table')return matrix(v);if(name==='COUNT'&&!Array.isArray(v)&&!isError(v)&&v!==null&&v!==''){try{return number(v);}catch{return v;}}return numericAggregate.has(name)&&!Array.isArray(v)?number(v):v;}));
    }
    if(fn&&scalarFunctions.has(name))return lifted(args.map(ev),fn);
    if(!fn&&(ctx.vars.has(name)||this.Workbook._names.has(name))){
      const lambda=ev({type:'name',name});
      if(isError(lambda))return lambda;
      if(lambda?.type==='lambda')return this._invoke(lambda,args.map(a=>a.omitted?OMITTED:ev(a)));
    }
    return super._evaluate(node,ctx);
  }
}
