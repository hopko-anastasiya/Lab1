using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lab1
{
    public class DependencyExtractor : FormulaBaseVisitor<object>
    {
        public HashSet<string> ReferencedCells { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        public override object VisitCellExpr(FormulaParser.CellExprContext context)
        {
            ReferencedCells.Add(context.GetText());
            return base.VisitCellExpr(context);
        }

        public override object VisitFormula(FormulaParser.FormulaContext context) => Visit(context.expr());
        public override object VisitNumberExpr(FormulaParser.NumberExprContext context) => null;
        public override object VisitAddExpr(FormulaParser.AddExprContext context) => base.VisitAddExpr(context);
        public override object VisitSubExpr(FormulaParser.SubExprContext context) => base.VisitSubExpr(context);
        public override object VisitMulExpr(FormulaParser.MulExprContext context) => base.VisitMulExpr(context);
        public override object VisitDivExpr(FormulaParser.DivExprContext context) => base.VisitDivExpr(context);
        public override object VisitPowerExpr(FormulaParser.PowerExprContext context) => base.VisitPowerExpr(context);
        public override object VisitUnaryMinusExpr(FormulaParser.UnaryMinusExprContext context) => base.VisitUnaryMinusExpr(context);
        public override object VisitUnaryPlusExpr(FormulaParser.UnaryPlusExprContext context) => base.VisitUnaryPlusExpr(context);
        public override object VisitParensExpr(FormulaParser.ParensExprContext context) => base.VisitParensExpr(context);
        public override object VisitEqualExpr(FormulaParser.EqualExprContext context) => base.VisitEqualExpr(context);
        public override object VisitNotEqualExpr(FormulaParser.NotEqualExprContext context) => base.VisitNotEqualExpr(context);
        public override object VisitLessExpr(FormulaParser.LessExprContext context) => base.VisitLessExpr(context);
        public override object VisitGreaterExpr(FormulaParser.GreaterExprContext context) => base.VisitGreaterExpr(context);
        public override object VisitLessEqualExpr(FormulaParser.LessEqualExprContext context) => base.VisitLessEqualExpr(context);
        public override object VisitGreaterEqualExpr(FormulaParser.GreaterEqualExprContext context) => base.VisitGreaterEqualExpr(context);
    }
}
