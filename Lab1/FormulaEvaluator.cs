using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Lab1.Data;

namespace Lab1
{
    public class FormulaEvaluator : FormulaBaseVisitor<double>
    {
        private readonly Dictionary<string, double> _cellValues;
        private readonly Table _table;

        public FormulaEvaluator(Dictionary<string, double> cellValues, Table table)
        {
            _cellValues = cellValues;
            _table = table;
        }

        public override double VisitFormula(FormulaParser.FormulaContext context)
        {
            return Visit(context.expr());
        }

        public override double VisitNumberExpr(FormulaParser.NumberExprContext context)
        {
            if (double.TryParse(context.GetText(), out double result))
                return result;
            throw new Exception($"Некоректне число: {context.GetText()}");
        }

        public override double VisitCellExpr(FormulaParser.CellExprContext context)
        {
            string cellName = context.GetText();
            if (_cellValues.TryGetValue(cellName, out double value))
                return value;
            Lab1.Data.Cell targetCell;
            try
            {
                targetCell = _table.GetCellByName(cellName);
                if (targetCell.Value?.ToString().StartsWith("#") == true)
                {
                    throw new InvalidOperationException($"Reference to error cell: {cellName}");
                }

                return 0;
            }
            catch (ArgumentException)
            {
                throw new InvalidOperationException($"Reference error: {cellName} does not exist.");
            }
            catch (InvalidOperationException)
            {
                throw;
            }
        }

        public override double VisitAddExpr(FormulaParser.AddExprContext context)
        {
            double left = Visit(context.expr(0));
            double right = Visit(context.expr(1));
            return left + right;
        }

        public override double VisitSubExpr(FormulaParser.SubExprContext context)
        {
            double left = Visit(context.expr(0));
            double right = Visit(context.expr(1));
            return left - right;
        }

        public override double VisitMulExpr(FormulaParser.MulExprContext context)
        {
            double left = Visit(context.expr(0));
            double right = Visit(context.expr(1));
            return left * right;
        }
        
        public override double VisitDivExpr(FormulaParser.DivExprContext context)
        {
            double left = Visit(context.expr(0));
            double right = Visit(context.expr(1));
            return left / right;
        }

        public override double VisitPowerExpr(FormulaParser.PowerExprContext context)
        {
            double left = Visit(context.expr(0));
            double right = Visit(context.expr(1));

            if (left == 0.0 && right == 0.0)
            {
                return double.NaN;
            }

            return Math.Pow(left, right);
        }

        public override double VisitUnaryMinusExpr(FormulaParser.UnaryMinusExprContext context)
        {
            return -Visit(context.expr());
        }

        public override double VisitUnaryPlusExpr(FormulaParser.UnaryPlusExprContext context)
        {
            return Visit(context.expr());
        }

        public override double VisitParensExpr(FormulaParser.ParensExprContext context)
        {
            return Visit(context.expr());
        }

        public override double VisitEqualExpr(FormulaParser.EqualExprContext context)
        {
            return Visit(context.expr(0)) == Visit(context.expr(1)) ? 1.0 : 0.0;
        }

        public override double VisitNotEqualExpr(FormulaParser.NotEqualExprContext context)
        {
            return Visit(context.expr(0)) != Visit(context.expr(1)) ? 1.0 : 0.0;
        }

        public override double VisitLessExpr(FormulaParser.LessExprContext context)
        {
            return Visit(context.expr(0)) < Visit(context.expr(1)) ? 1.0 : 0.0;
        }

        public override double VisitGreaterExpr(FormulaParser.GreaterExprContext context)
        {
            return Visit(context.expr(0)) > Visit(context.expr(1)) ? 1.0 : 0.0;
        }

        public override double VisitLessEqualExpr(FormulaParser.LessEqualExprContext context)
        {
            return Visit(context.expr(0)) <= Visit(context.expr(1)) ? 1.0 : 0.0;
        }

        public override double VisitGreaterEqualExpr(FormulaParser.GreaterEqualExprContext context)
        {
            return Visit(context.expr(0)) >= Visit(context.expr(1)) ? 1.0 : 0.0;
        }
    }
}
