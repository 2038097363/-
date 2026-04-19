using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace MatrixWorkspaceApp
{
    internal enum ValueKind
    {
        Matrix,
        Scalar,
        Text
    }

    internal sealed class EvalValue
    {
        public ValueKind Kind { get; private set; }
        public Matrix MatrixValue { get; private set; }
        public double ScalarValue { get; private set; }
        public string TextValue { get; private set; }

        public static EvalValue FromMatrix(Matrix matrix)
        {
            return new EvalValue { Kind = ValueKind.Matrix, MatrixValue = matrix };
        }

        public static EvalValue FromScalar(double value)
        {
            return new EvalValue { Kind = ValueKind.Scalar, ScalarValue = value };
        }

        public static EvalValue FromText(string value)
        {
            return new EvalValue { Kind = ValueKind.Text, TextValue = value };
        }
    }

    internal sealed class Matrix
    {
        private readonly double[,] _data;

        public int Rows { get; private set; }
        public int Cols { get; private set; }

        public Matrix(int rows, int cols)
        {
            Rows = rows;
            Cols = cols;
            _data = new double[rows, cols];
        }

        public double this[int row, int col]
        {
            get { return _data[row, col]; }
            set { _data[row, col] = value; }
        }

        public Matrix Clone()
        {
            var clone = new Matrix(Rows, Cols);
            for (int r = 0; r < Rows; r++)
            {
                for (int c = 0; c < Cols; c++)
                {
                    clone[r, c] = this[r, c];
                }
            }

            return clone;
        }

        public static Matrix Zero(int rows, int cols)
        {
            return new Matrix(rows, cols);
        }

        public static Matrix Identity(int size)
        {
            var matrix = new Matrix(size, size);
            for (int i = 0; i < size; i++)
            {
                matrix[i, i] = 1.0;
            }

            return matrix;
        }
    }

    internal sealed class MatrixResult
    {
        public Matrix Matrix { get; set; }
        public int Rank { get; set; }
        public List<int> PivotColumns { get; set; }
    }

    internal static class UiText
    {
        public const string Title = "\u77e9\u9635\u5de5\u4f5c\u53f0";
        public const string Subtitle = "\u591a\u77e9\u9635\u8f93\u5165\u4e0e\u8868\u8fbe\u5f0f\u8ba1\u7b97";
        public const string MatrixCount = "\u77e9\u9635\u6570\u91cf";
        public const string AddMatrix = "\u589e\u52a0";
        public const string RemoveMatrix = "\u51cf\u5c11";
        public const string OpenExpression = "\u8868\u8fbe\u5f0f\u7a97\u53e3";
        public const string OpenResult = "\u7ed3\u679c\u7a97\u53e3";
        public const string ExpressionWindow = "\u8868\u8fbe\u5f0f\u7a97\u53e3";
        public const string ResultWindow = "\u7ed3\u679c\u7a97\u53e3";
        public const string Expression = "\u8868\u8fbe\u5f0f";
        public const string Run = "\u8ba1\u7b97\u5e76\u663e\u793a\u7ed3\u679c";
        public const string Sample = "\u793a\u4f8b";
        public const string Clear = "\u6e05\u7a7a";
        public const string Ready = "\u5c31\u7eea";
        public const string Done = "\u8ba1\u7b97\u5b8c\u6210";
        public const string EvalFailed = "\u8ba1\u7b97\u5931\u8d25";
        public const string Result = "\u7ed3\u679c";
        public const string InputTip = "[[1,2],[3,4]]  or  [1 2; 3 4]";
        public const string ExprHint = "\u8f93\u5165\u5b8c\u8868\u8fbe\u5f0f\u540e\uff0c\u70b9\u51fb\u201c\u8ba1\u7b97\u5e76\u663e\u793a\u7ed3\u679c\u201d\u6216\u76f4\u63a5\u6309\u56de\u8f66\u3002";
        public const string Syntax = "A*B, A/B, A^2, RREF(A), REF(A), TR(A), NORM(A), ADJ(A), COF(A), INV(A), DET(B), NULL(A)";
        public const string ParseEmptySuffix = "\u4e0d\u80fd\u4e3a\u7a7a\u3002";
        public const string ParseInvalidSuffix = "\u8f93\u5165\u683c\u5f0f\u65e0\u6548\u3002";
        public const string ParseEmptyMatrixSuffix = "\u5fc5\u987b\u662f\u975e\u7a7a\u4e8c\u7ef4\u6570\u7ec4\u3002";
        public const string ParseEmptyRowSuffix = "\u7684\u6bcf\u4e00\u884c\u90fd\u5fc5\u987b\u662f\u975e\u7a7a\u6570\u7ec4\u3002";
        public const string ParseRowLenSuffix = "\u7684\u6bcf\u4e00\u884c\u5217\u6570\u5fc5\u987b\u4e00\u81f4\u3002";
        public const string ParseNumericSuffix = "\u4e2d\u53ea\u80fd\u5305\u542b\u6570\u5b57\u3002";
        public const string UnknownVar = "\u672a\u627e\u5230\u77e9\u9635\u53d8\u91cf\uff1a";
        public const string NeedMatrix = "\u8be5\u8fd0\u7b97\u9700\u8981\u77e9\u9635\u53c2\u6570\u3002";
        public const string AddDim = "\u52a0\u6cd5\u8981\u6c42\u4e24\u4e2a\u77e9\u9635\u7ef4\u5ea6\u76f8\u540c\u3002";
        public const string SubDim = "\u51cf\u6cd5\u8981\u6c42\u4e24\u4e2a\u77e9\u9635\u7ef4\u5ea6\u76f8\u540c\u3002";
        public const string MulDim = "\u4e58\u6cd5\u8981\u6c42\u5de6\u77e9\u9635\u5217\u6570\u7b49\u4e8e\u53f3\u77e9\u9635\u884c\u6570\u3002";
        public const string DivZero = "\u9664\u6570\u4e0d\u80fd\u4e3a\u96f6\u3002";
        public const string SquareOnly = "\u53ea\u6709\u65b9\u9635\u624d\u80fd\u8fdb\u884c\u8be5\u8fd0\u7b97\u3002";
        public const string DetSquare = "\u53ea\u6709\u65b9\u9635\u624d\u80fd\u6c42\u884c\u5217\u5f0f\u3002";
        public const string InvSquare = "\u53ea\u6709\u65b9\u9635\u624d\u80fd\u6c42\u9006\u77e9\u9635\u3002";
        public const string IntegerPowerOnly = "\u77e9\u9635\u5e42\u8fd0\u7b97\u53ea\u652f\u6301\u6574\u6570\u6b21\u5e42\u3002";
        public const string NotInvertible = "\u8be5\u77e9\u9635\u4e0d\u53ef\u9006\u3002";
        public const string UnexpectedToken = "\u8868\u8fbe\u5f0f\u542b\u6709\u65e0\u6548\u7b26\u53f7\u3002";
        public const string MissingParen = "\u7f3a\u5c11\u53f3\u62ec\u53f7\u3002";
        public const string UnsupportedOp = "\u4e0d\u652f\u6301\u8be5\u8fd0\u7b97\u7ec4\u5408\u3002";
        public const string TrivialNull = "\u96f6\u7a7a\u95f4\uff1a\u53ea\u542b\u96f6\u5411\u91cf";
        public const string MaxMatrices = "\u6700\u591a\u652f\u6301 6 \u4e2a\u77e9\u9635\u53d8\u91cf\u3002";
    }

    internal static class NumberUtil
    {
        public static bool IsZero(double value)
        {
            return Math.Abs(value) < 1e-10;
        }

        public static string Format(double value)
        {
            if (IsZero(value))
            {
                value = 0.0;
            }

            if (Math.Abs(value - Math.Round(value)) < 1e-10)
            {
                return Math.Round(value).ToString(CultureInfo.InvariantCulture);
            }

            return value.ToString("0.######", CultureInfo.InvariantCulture);
        }
    }

    internal static class MatrixFormatter
    {
        public static string ToMatlab(Matrix matrix)
        {
            if (matrix == null || matrix.Rows == 0 || matrix.Cols == 0)
            {
                return "[]";
            }

            var lines = new List<string>();
            for (int r = 0; r < matrix.Rows; r++)
            {
                var cells = new List<string>();
                for (int c = 0; c < matrix.Cols; c++)
                {
                    cells.Add(NumberUtil.Format(matrix[r, c]));
                }

                string rowText = string.Join("  ", cells);
                if (matrix.Rows == 1)
                {
                    lines.Add("[" + rowText + "]");
                }
                else if (r == 0)
                {
                    lines.Add("[" + rowText);
                }
                else if (r == matrix.Rows - 1)
                {
                    lines.Add(" " + rowText + "]");
                }
                else
                {
                    lines.Add(" " + rowText);
                }
            }

            return string.Join(Environment.NewLine, lines);
        }
    }

    internal static class MatrixInputParser
    {
        public static Matrix Parse(string text, string name)
        {
            string trimmed = Normalize(text);
            if (string.IsNullOrWhiteSpace(trimmed))
            {
                throw new InvalidOperationException(name + UiText.ParseEmptySuffix);
            }

            if (!trimmed.StartsWith("[", StringComparison.Ordinal) || !trimmed.EndsWith("]", StringComparison.Ordinal))
            {
                throw new InvalidOperationException(name + UiText.ParseInvalidSuffix);
            }

            string inner = trimmed.Substring(1, trimmed.Length - 2).Trim();
            if (string.IsNullOrWhiteSpace(inner))
            {
                throw new InvalidOperationException(name + UiText.ParseEmptyMatrixSuffix);
            }

            string[] rowTexts = inner.IndexOf("][", StringComparison.Ordinal) >= 0
                ? Regex.Split(inner, @"\]\s*,\s*\[")
                : inner.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);

            var rows = new List<double[]>();
            int? cols = null;

            foreach (string rawRow in rowTexts)
            {
                string rowText = rawRow.Trim().Trim('[', ']');
                rowText = rowText.Replace(",", " ");
                string[] parts = rowText.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length == 0)
                {
                    throw new InvalidOperationException(name + UiText.ParseEmptyRowSuffix);
                }

                if (!cols.HasValue)
                {
                    cols = parts.Length;
                }
                else if (cols.Value != parts.Length)
                {
                    throw new InvalidOperationException(name + UiText.ParseRowLenSuffix);
                }

                var row = new double[parts.Length];
                for (int i = 0; i < parts.Length; i++)
                {
                    double value;
                    if (!double.TryParse(parts[i], NumberStyles.Float, CultureInfo.InvariantCulture, out value) &&
                        !double.TryParse(parts[i], NumberStyles.Float, CultureInfo.CurrentCulture, out value))
                    {
                        throw new InvalidOperationException(name + UiText.ParseNumericSuffix);
                    }

                    row[i] = value;
                }

                rows.Add(row);
            }

            var matrix = new Matrix(rows.Count, cols ?? 0);
            for (int r = 0; r < rows.Count; r++)
            {
                for (int c = 0; c < rows[r].Length; c++)
                {
                    matrix[r, c] = rows[r][c];
                }
            }

            return matrix;
        }

        private static string Normalize(string text)
        {
            string trimmed = (text ?? string.Empty).Trim();
            if (trimmed.StartsWith("[[", StringComparison.Ordinal))
            {
                return trimmed;
            }

            return trimmed;
        }
    }

    internal static class MatrixMath
    {
        public static Matrix Add(Matrix a, Matrix b)
        {
            EnsureSameSize(a, b, UiText.AddDim);
            var result = Matrix.Zero(a.Rows, a.Cols);
            for (int r = 0; r < a.Rows; r++)
            {
                for (int c = 0; c < a.Cols; c++)
                {
                    result[r, c] = a[r, c] + b[r, c];
                }
            }

            return result;
        }

        public static Matrix Subtract(Matrix a, Matrix b)
        {
            EnsureSameSize(a, b, UiText.SubDim);
            var result = Matrix.Zero(a.Rows, a.Cols);
            for (int r = 0; r < a.Rows; r++)
            {
                for (int c = 0; c < a.Cols; c++)
                {
                    result[r, c] = a[r, c] - b[r, c];
                }
            }

            return result;
        }

        public static Matrix Multiply(Matrix a, Matrix b)
        {
            if (a.Cols != b.Rows)
            {
                throw new InvalidOperationException(UiText.MulDim);
            }

            var result = Matrix.Zero(a.Rows, b.Cols);
            for (int r = 0; r < a.Rows; r++)
            {
                for (int c = 0; c < b.Cols; c++)
                {
                    double sum = 0.0;
                    for (int k = 0; k < a.Cols; k++)
                    {
                        sum += a[r, k] * b[k, c];
                    }

                    result[r, c] = sum;
                }
            }

            return result;
        }

        public static Matrix Multiply(Matrix matrix, double scalar)
        {
            var result = Matrix.Zero(matrix.Rows, matrix.Cols);
            for (int r = 0; r < matrix.Rows; r++)
            {
                for (int c = 0; c < matrix.Cols; c++)
                {
                    result[r, c] = matrix[r, c] * scalar;
                }
            }

            return result;
        }

        public static Matrix Divide(Matrix matrix, double scalar)
        {
            if (NumberUtil.IsZero(scalar))
            {
                throw new InvalidOperationException(UiText.DivZero);
            }

            return Multiply(matrix, 1.0 / scalar);
        }

        public static Matrix Transpose(Matrix matrix)
        {
            var result = Matrix.Zero(matrix.Cols, matrix.Rows);
            for (int r = 0; r < matrix.Rows; r++)
            {
                for (int c = 0; c < matrix.Cols; c++)
                {
                    result[c, r] = matrix[r, c];
                }
            }

            return result;
        }

        public static MatrixResult Rref(Matrix matrix)
        {
            var working = matrix.Clone();
            int lead = 0;
            int rank = 0;
            var pivotColumns = new List<int>();

            for (int r = 0; r < working.Rows; r++)
            {
                if (lead >= working.Cols)
                {
                    break;
                }

                int pivotRow = r;
                while (pivotRow < working.Rows && NumberUtil.IsZero(working[pivotRow, lead]))
                {
                    pivotRow++;
                }

                while (pivotRow >= working.Rows)
                {
                    lead++;
                    if (lead >= working.Cols)
                    {
                        return new MatrixResult { Matrix = working, Rank = rank, PivotColumns = pivotColumns };
                    }

                    pivotRow = r;
                    while (pivotRow < working.Rows && NumberUtil.IsZero(working[pivotRow, lead]))
                    {
                        pivotRow++;
                    }
                }

                if (pivotRow != r)
                {
                    SwapRows(working, r, pivotRow);
                }

                double pivot = working[r, lead];
                for (int c = 0; c < working.Cols; c++)
                {
                    working[r, c] /= pivot;
                }

                for (int row = 0; row < working.Rows; row++)
                {
                    if (row == r)
                    {
                        continue;
                    }

                    double factor = working[row, lead];
                    if (NumberUtil.IsZero(factor))
                    {
                        continue;
                    }

                    for (int c = 0; c < working.Cols; c++)
                    {
                        working[row, c] -= factor * working[r, c];
                        if (NumberUtil.IsZero(working[row, c]))
                        {
                            working[row, c] = 0.0;
                        }
                    }
                }

                pivotColumns.Add(lead);
                rank++;
                lead++;
            }

            return new MatrixResult { Matrix = working, Rank = rank, PivotColumns = pivotColumns };
        }

        public static Matrix Ref(Matrix matrix)
        {
            var working = matrix.Clone();
            int pivotRow = 0;

            for (int pivotCol = 0; pivotCol < working.Cols && pivotRow < working.Rows; pivotCol++)
            {
                int best = pivotRow;
                while (best < working.Rows && NumberUtil.IsZero(working[best, pivotCol]))
                {
                    best++;
                }

                if (best >= working.Rows)
                {
                    continue;
                }

                if (best != pivotRow)
                {
                    SwapRows(working, best, pivotRow);
                }

                for (int row = pivotRow + 1; row < working.Rows; row++)
                {
                    double factor = working[row, pivotCol] / working[pivotRow, pivotCol];
                    if (NumberUtil.IsZero(factor))
                    {
                        continue;
                    }

                    for (int c = pivotCol; c < working.Cols; c++)
                    {
                        working[row, c] -= factor * working[pivotRow, c];
                        if (NumberUtil.IsZero(working[row, c]))
                        {
                            working[row, c] = 0.0;
                        }
                    }
                }

                pivotRow++;
            }

            return working;
        }

        public static int Rank(Matrix matrix)
        {
            return Rref(matrix).Rank;
        }

        public static double Trace(Matrix matrix)
        {
            EnsureSquare(matrix, UiText.SquareOnly);
            double sum = 0.0;
            for (int i = 0; i < matrix.Rows; i++)
            {
                sum += matrix[i, i];
            }

            return sum;
        }

        public static double Norm(Matrix matrix)
        {
            double sum = 0.0;
            for (int r = 0; r < matrix.Rows; r++)
            {
                for (int c = 0; c < matrix.Cols; c++)
                {
                    sum += matrix[r, c] * matrix[r, c];
                }
            }

            return Math.Sqrt(sum);
        }

        public static double Determinant(Matrix matrix)
        {
            EnsureSquare(matrix, UiText.DetSquare);
            var working = matrix.Clone();
            double det = 1.0;
            int sign = 1;

            for (int col = 0; col < working.Cols; col++)
            {
                int pivot = col;
                while (pivot < working.Rows && NumberUtil.IsZero(working[pivot, col]))
                {
                    pivot++;
                }

                if (pivot >= working.Rows)
                {
                    return 0.0;
                }

                if (pivot != col)
                {
                    SwapRows(working, pivot, col);
                    sign *= -1;
                }

                double pivotValue = working[col, col];
                det *= pivotValue;

                for (int row = col + 1; row < working.Rows; row++)
                {
                    double factor = working[row, col] / pivotValue;
                    for (int c = col; c < working.Cols; c++)
                    {
                        working[row, c] -= factor * working[col, c];
                    }
                }
            }

            return det * sign;
        }

        public static Matrix Inverse(Matrix matrix)
        {
            EnsureSquare(matrix, UiText.InvSquare);
            int size = matrix.Rows;
            var left = matrix.Clone();
            var right = Matrix.Identity(size);

            for (int col = 0; col < size; col++)
            {
                int pivot = col;
                while (pivot < size && NumberUtil.IsZero(left[pivot, col]))
                {
                    pivot++;
                }

                if (pivot >= size)
                {
                    throw new InvalidOperationException(UiText.NotInvertible);
                }

                if (pivot != col)
                {
                    SwapRows(left, pivot, col);
                    SwapRows(right, pivot, col);
                }

                double pivotValue = left[col, col];
                if (NumberUtil.IsZero(pivotValue))
                {
                    throw new InvalidOperationException(UiText.NotInvertible);
                }

                for (int c = 0; c < size; c++)
                {
                    left[col, c] /= pivotValue;
                    right[col, c] /= pivotValue;
                }

                for (int row = 0; row < size; row++)
                {
                    if (row == col)
                    {
                        continue;
                    }

                    double factor = left[row, col];
                    if (NumberUtil.IsZero(factor))
                    {
                        continue;
                    }

                    for (int c = 0; c < size; c++)
                    {
                        left[row, c] -= factor * left[col, c];
                        right[row, c] -= factor * right[col, c];
                    }
                }
            }

            return right;
        }

        public static Matrix Cofactor(Matrix matrix)
        {
            EnsureSquare(matrix, UiText.SquareOnly);
            var result = Matrix.Zero(matrix.Rows, matrix.Cols);
            for (int r = 0; r < matrix.Rows; r++)
            {
                for (int c = 0; c < matrix.Cols; c++)
                {
                    double minor = Determinant(GetMinor(matrix, r, c));
                    result[r, c] = ((r + c) % 2 == 0 ? 1 : -1) * minor;
                }
            }

            return result;
        }

        public static Matrix Adjoint(Matrix matrix)
        {
            return Transpose(Cofactor(matrix));
        }

        public static Matrix Power(Matrix matrix, int exponent)
        {
            EnsureSquare(matrix, UiText.SquareOnly);
            if (exponent == 0)
            {
                return Matrix.Identity(matrix.Rows);
            }

            if (exponent < 0)
            {
                return Power(Inverse(matrix), -exponent);
            }

            var result = Matrix.Identity(matrix.Rows);
            var baseMatrix = matrix.Clone();
            int power = exponent;
            while (power > 0)
            {
                if ((power & 1) == 1)
                {
                    result = Multiply(result, baseMatrix);
                }

                power >>= 1;
                if (power > 0)
                {
                    baseMatrix = Multiply(baseMatrix, baseMatrix);
                }
            }

            return result;
        }

        public static string NullSpace(Matrix matrix)
        {
            MatrixResult rref = Rref(matrix);
            var pivots = new HashSet<int>(rref.PivotColumns);
            var freeColumns = Enumerable.Range(0, matrix.Cols).Where(index => !pivots.Contains(index)).ToList();
            if (freeColumns.Count == 0)
            {
                return UiText.TrivialNull;
            }

            var basis = Matrix.Zero(matrix.Cols, freeColumns.Count);
            for (int freeIndex = 0; freeIndex < freeColumns.Count; freeIndex++)
            {
                int freeColumn = freeColumns[freeIndex];
                basis[freeColumn, freeIndex] = 1.0;
                for (int row = 0; row < rref.PivotColumns.Count; row++)
                {
                    int pivotColumn = rref.PivotColumns[row];
                    basis[pivotColumn, freeIndex] = -rref.Matrix[row, freeColumn];
                }
            }

            return MatrixFormatter.ToMatlab(basis);
        }

        private static Matrix GetMinor(Matrix matrix, int skipRow, int skipCol)
        {
            var result = Matrix.Zero(matrix.Rows - 1, matrix.Cols - 1);
            int rr = 0;
            for (int r = 0; r < matrix.Rows; r++)
            {
                if (r == skipRow)
                {
                    continue;
                }

                int cc = 0;
                for (int c = 0; c < matrix.Cols; c++)
                {
                    if (c == skipCol)
                    {
                        continue;
                    }

                    result[rr, cc] = matrix[r, c];
                    cc++;
                }

                rr++;
            }

            return result;
        }

        private static void EnsureSameSize(Matrix a, Matrix b, string message)
        {
            if (a.Rows != b.Rows || a.Cols != b.Cols)
            {
                throw new InvalidOperationException(message);
            }
        }

        private static void EnsureSquare(Matrix matrix, string message)
        {
            if (matrix.Rows != matrix.Cols)
            {
                throw new InvalidOperationException(message);
            }
        }

        private static void SwapRows(Matrix matrix, int rowA, int rowB)
        {
            for (int c = 0; c < matrix.Cols; c++)
            {
                double temp = matrix[rowA, c];
                matrix[rowA, c] = matrix[rowB, c];
                matrix[rowB, c] = temp;
            }
        }
    }

    internal enum TokenType
    {
        Identifier,
        Number,
        Plus,
        Minus,
        Star,
        Slash,
        Caret,
        LeftParen,
        RightParen,
        Comma,
        End
    }

    internal sealed class Token
    {
        public Token(TokenType type, string text)
        {
            Type = type;
            Text = text;
        }

        public TokenType Type { get; private set; }
        public string Text { get; private set; }
    }

    internal sealed class ExpressionTokenizer
    {
        private readonly string _text;
        private int _index;

        public ExpressionTokenizer(string text)
        {
            _text = text ?? string.Empty;
        }

        public List<Token> Tokenize()
        {
            var tokens = new List<Token>();
            while (_index < _text.Length)
            {
                char ch = _text[_index];
                if (char.IsWhiteSpace(ch))
                {
                    _index++;
                    continue;
                }

                if (char.IsLetter(ch))
                {
                    int start = _index;
                    while (_index < _text.Length && (char.IsLetterOrDigit(_text[_index]) || _text[_index] == '_'))
                    {
                        _index++;
                    }

                    tokens.Add(new Token(TokenType.Identifier, _text.Substring(start, _index - start)));
                    continue;
                }

                if (char.IsDigit(ch) || ch == '.')
                {
                    int start = _index;
                    _index++;
                    while (_index < _text.Length && (char.IsDigit(_text[_index]) || _text[_index] == '.'))
                    {
                        _index++;
                    }

                    tokens.Add(new Token(TokenType.Number, _text.Substring(start, _index - start)));
                    continue;
                }

                _index++;
                switch (ch)
                {
                    case '+':
                        tokens.Add(new Token(TokenType.Plus, "+"));
                        break;
                    case '-':
                        tokens.Add(new Token(TokenType.Minus, "-"));
                        break;
                    case '*':
                        tokens.Add(new Token(TokenType.Star, "*"));
                        break;
                    case '/':
                        tokens.Add(new Token(TokenType.Slash, "/"));
                        break;
                    case '^':
                        tokens.Add(new Token(TokenType.Caret, "^"));
                        break;
                    case '(':
                        tokens.Add(new Token(TokenType.LeftParen, "("));
                        break;
                    case ')':
                        tokens.Add(new Token(TokenType.RightParen, ")"));
                        break;
                    case ',':
                        tokens.Add(new Token(TokenType.Comma, ","));
                        break;
                    default:
                        throw new InvalidOperationException(UiText.UnexpectedToken);
                }
            }

            tokens.Add(new Token(TokenType.End, string.Empty));
            return tokens;
        }
    }

    internal sealed class ExpressionParser
    {
        private readonly List<Token> _tokens;
        private readonly Dictionary<string, Matrix> _context;
        private int _index;

        public ExpressionParser(string expression, Dictionary<string, Matrix> context)
        {
            _tokens = new ExpressionTokenizer(expression).Tokenize();
            _context = context;
        }

        public EvalValue Parse()
        {
            EvalValue value = ParseExpression();
            if (Current.Type != TokenType.End)
            {
                throw new InvalidOperationException(UiText.UnexpectedToken);
            }

            return value;
        }

        private EvalValue ParseExpression()
        {
            EvalValue value = ParseTerm();
            while (Current.Type == TokenType.Plus || Current.Type == TokenType.Minus)
            {
                TokenType op = Current.Type;
                Next();
                EvalValue right = ParseTerm();
                value = ApplyBinary(op, value, right);
            }

            return value;
        }

        private EvalValue ParseTerm()
        {
            EvalValue value = ParsePower();
            while (Current.Type == TokenType.Star || Current.Type == TokenType.Slash)
            {
                TokenType op = Current.Type;
                Next();
                EvalValue right = ParsePower();
                value = ApplyBinary(op, value, right);
            }

            return value;
        }

        private EvalValue ParsePower()
        {
            EvalValue value = ParseUnary();
            while (Current.Type == TokenType.Caret)
            {
                Next();
                EvalValue right = ParseUnary();
                value = ApplyBinary(TokenType.Caret, value, right);
            }

            return value;
        }

        private EvalValue ParseUnary()
        {
            if (Current.Type == TokenType.Minus)
            {
                Next();
                EvalValue value = ParseUnary();
                if (value.Kind == ValueKind.Scalar)
                {
                    return EvalValue.FromScalar(-value.ScalarValue);
                }

                if (value.Kind == ValueKind.Matrix)
                {
                    return EvalValue.FromMatrix(MatrixMath.Multiply(value.MatrixValue, -1.0));
                }

                throw new InvalidOperationException(UiText.UnsupportedOp);
            }

            return ParsePrimary();
        }

        private EvalValue ParsePrimary()
        {
            if (Current.Type == TokenType.Number)
            {
                double number;
                if (!double.TryParse(Current.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out number))
                {
                    throw new InvalidOperationException(UiText.UnexpectedToken);
                }

                Next();
                return EvalValue.FromScalar(number);
            }

            if (Current.Type == TokenType.Identifier)
            {
                string identifier = Current.Text;
                Next();

                if (Current.Type == TokenType.LeftParen)
                {
                    return ParseFunction(identifier);
                }

                Matrix matrix;
                if (!_context.TryGetValue(identifier.ToUpperInvariant(), out matrix))
                {
                    throw new InvalidOperationException(UiText.UnknownVar + identifier);
                }

                return EvalValue.FromMatrix(matrix);
            }

            if (Current.Type == TokenType.LeftParen)
            {
                Next();
                EvalValue value = ParseExpression();
                Expect(TokenType.RightParen, UiText.MissingParen);
                return value;
            }

            throw new InvalidOperationException(UiText.UnexpectedToken);
        }

        private EvalValue ParseFunction(string name)
        {
            Expect(TokenType.LeftParen, UiText.UnexpectedToken);
            var args = new List<EvalValue>();
            if (Current.Type != TokenType.RightParen)
            {
                while (true)
                {
                    args.Add(ParseExpression());
                    if (Current.Type != TokenType.Comma)
                    {
                        break;
                    }

                    Next();
                }
            }

            Expect(TokenType.RightParen, UiText.MissingParen);
            string upper = name.ToUpperInvariant();

            switch (upper)
            {
                case "T":
                    return EvalValue.FromMatrix(MatrixMath.Transpose(RequireMatrix(args, 1)[0]));
                case "REF":
                    return EvalValue.FromMatrix(MatrixMath.Ref(RequireMatrix(args, 1)[0]));
                case "RREF":
                    return EvalValue.FromMatrix(MatrixMath.Rref(RequireMatrix(args, 1)[0]).Matrix);
                case "RANK":
                    return EvalValue.FromScalar(MatrixMath.Rank(RequireMatrix(args, 1)[0]));
                case "TR":
                    return EvalValue.FromScalar(MatrixMath.Trace(RequireMatrix(args, 1)[0]));
                case "NORM":
                    return EvalValue.FromScalar(MatrixMath.Norm(RequireMatrix(args, 1)[0]));
                case "INV":
                    return EvalValue.FromMatrix(MatrixMath.Inverse(RequireMatrix(args, 1)[0]));
                case "DET":
                    return EvalValue.FromScalar(MatrixMath.Determinant(RequireMatrix(args, 1)[0]));
                case "COF":
                    return EvalValue.FromMatrix(MatrixMath.Cofactor(RequireMatrix(args, 1)[0]));
                case "ADJ":
                    return EvalValue.FromMatrix(MatrixMath.Adjoint(RequireMatrix(args, 1)[0]));
                case "NULL":
                    return EvalValue.FromText(MatrixMath.NullSpace(RequireMatrix(args, 1)[0]));
                default:
                    throw new InvalidOperationException(UiText.UnexpectedToken);
            }
        }

        private Matrix[] RequireMatrix(List<EvalValue> args, int count)
        {
            if (args.Count != count)
            {
                throw new InvalidOperationException(UiText.UnsupportedOp);
            }

            var matrices = new Matrix[count];
            for (int i = 0; i < count; i++)
            {
                if (args[i].Kind != ValueKind.Matrix)
                {
                    throw new InvalidOperationException(UiText.NeedMatrix);
                }

                matrices[i] = args[i].MatrixValue;
            }

            return matrices;
        }

        private EvalValue ApplyBinary(TokenType op, EvalValue left, EvalValue right)
        {
            switch (op)
            {
                case TokenType.Plus:
                    if (left.Kind == ValueKind.Matrix && right.Kind == ValueKind.Matrix)
                    {
                        return EvalValue.FromMatrix(MatrixMath.Add(left.MatrixValue, right.MatrixValue));
                    }

                    if (left.Kind == ValueKind.Scalar && right.Kind == ValueKind.Scalar)
                    {
                        return EvalValue.FromScalar(left.ScalarValue + right.ScalarValue);
                    }

                    break;
                case TokenType.Minus:
                    if (left.Kind == ValueKind.Matrix && right.Kind == ValueKind.Matrix)
                    {
                        return EvalValue.FromMatrix(MatrixMath.Subtract(left.MatrixValue, right.MatrixValue));
                    }

                    if (left.Kind == ValueKind.Scalar && right.Kind == ValueKind.Scalar)
                    {
                        return EvalValue.FromScalar(left.ScalarValue - right.ScalarValue);
                    }

                    break;
                case TokenType.Star:
                    if (left.Kind == ValueKind.Matrix && right.Kind == ValueKind.Matrix)
                    {
                        return EvalValue.FromMatrix(MatrixMath.Multiply(left.MatrixValue, right.MatrixValue));
                    }

                    if (left.Kind == ValueKind.Matrix && right.Kind == ValueKind.Scalar)
                    {
                        return EvalValue.FromMatrix(MatrixMath.Multiply(left.MatrixValue, right.ScalarValue));
                    }

                    if (left.Kind == ValueKind.Scalar && right.Kind == ValueKind.Matrix)
                    {
                        return EvalValue.FromMatrix(MatrixMath.Multiply(right.MatrixValue, left.ScalarValue));
                    }

                    if (left.Kind == ValueKind.Scalar && right.Kind == ValueKind.Scalar)
                    {
                        return EvalValue.FromScalar(left.ScalarValue * right.ScalarValue);
                    }

                    break;
                case TokenType.Slash:
                    if (left.Kind == ValueKind.Matrix && right.Kind == ValueKind.Scalar)
                    {
                        return EvalValue.FromMatrix(MatrixMath.Divide(left.MatrixValue, right.ScalarValue));
                    }

                    if (left.Kind == ValueKind.Matrix && right.Kind == ValueKind.Matrix)
                    {
                        return EvalValue.FromMatrix(MatrixMath.Multiply(left.MatrixValue, MatrixMath.Inverse(right.MatrixValue)));
                    }

                    if (left.Kind == ValueKind.Scalar && right.Kind == ValueKind.Scalar)
                    {
                        if (NumberUtil.IsZero(right.ScalarValue))
                        {
                            throw new InvalidOperationException(UiText.DivZero);
                        }

                        return EvalValue.FromScalar(left.ScalarValue / right.ScalarValue);
                    }

                    break;
                case TokenType.Caret:
                    if (left.Kind == ValueKind.Matrix && right.Kind == ValueKind.Scalar)
                    {
                        double rounded = Math.Round(right.ScalarValue);
                        if (Math.Abs(right.ScalarValue - rounded) > 1e-10)
                        {
                            throw new InvalidOperationException(UiText.IntegerPowerOnly);
                        }

                        return EvalValue.FromMatrix(MatrixMath.Power(left.MatrixValue, (int)rounded));
                    }

                    if (left.Kind == ValueKind.Scalar && right.Kind == ValueKind.Scalar)
                    {
                        return EvalValue.FromScalar(Math.Pow(left.ScalarValue, right.ScalarValue));
                    }

                    break;
            }

            throw new InvalidOperationException(UiText.UnsupportedOp);
        }

        private Token Current
        {
            get { return _tokens[_index]; }
        }

        private void Next()
        {
            if (_index < _tokens.Count - 1)
            {
                _index++;
            }
        }

        private void Expect(TokenType type, string message)
        {
            if (Current.Type != type)
            {
                throw new InvalidOperationException(message);
            }

            Next();
        }
    }

    internal sealed class ResultForm : Form
    {
        private readonly Label _statusLabel;
        private readonly TextBox _resultBox;

        public ResultForm()
        {
            Text = UiText.ResultWindow;
            StartPosition = FormStartPosition.Manual;
            Size = new Size(620, 420);
            MinimumSize = new Size(500, 320);
            BackColor = Color.FromArgb(247, 250, 252);

            var title = new Label
            {
                Text = UiText.Result,
                Font = new Font("Microsoft YaHei UI", 15f, FontStyle.Bold),
                ForeColor = Color.FromArgb(24, 54, 96),
                AutoSize = true,
                Location = new Point(20, 18)
            };

            _statusLabel = new Label
            {
                Text = UiText.Ready,
                Font = new Font("Microsoft YaHei UI", 9f),
                ForeColor = Color.FromArgb(88, 111, 145),
                AutoSize = true,
                Location = new Point(24, 52)
            };

            _resultBox = new TextBox
            {
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = true,
                Font = new Font("Consolas", 12.5f),
                BorderStyle = BorderStyle.FixedSingle,
                Location = new Point(20, 78),
                Size = new Size(564, 284),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };

            Controls.Add(title);
            Controls.Add(_statusLabel);
            Controls.Add(_resultBox);
        }

        public void ShowResult(string status, string result)
        {
            _statusLabel.Text = status;
            _resultBox.Text = result;
            if (!Visible)
            {
                Show();
            }

            BringToFront();
            Activate();
        }
    }

    internal sealed class ExpressionForm : Form
    {
        private readonly MainForm _mainForm;
        private readonly TextBox _expressionBox;
        private bool _allowClose;

        public ExpressionForm(MainForm mainForm)
        {
            _mainForm = mainForm;
            Text = UiText.ExpressionWindow;
            StartPosition = FormStartPosition.Manual;
            Size = new Size(980, 260);
            MinimumSize = new Size(760, 220);
            BackColor = Color.White;

            var title = new Label
            {
                Text = UiText.Expression,
                Font = new Font("Microsoft YaHei UI", 16f, FontStyle.Bold),
                ForeColor = Color.FromArgb(20, 44, 74),
                AutoSize = true,
                Location = new Point(20, 20)
            };

            _expressionBox = new TextBox
            {
                Font = new Font("Consolas", 18f),
                Location = new Point(138, 16),
                Size = new Size(555, 40),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };

            var runButton = BuildActionButton(UiText.Run, Color.FromArgb(41, 131, 156), Color.White, new Point(708, 14), new Size(155, 42));
            runButton.Click += delegate { _mainForm.ExecuteExpression(_expressionBox.Text); };
            AcceptButton = runButton;

            var sampleButton = BuildActionButton(UiText.Sample, Color.FromArgb(246, 236, 213), Color.FromArgb(133, 86, 8), new Point(874, 14), new Size(80, 42));
            sampleButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            sampleButton.Click += delegate { _expressionBox.Text = "A*B"; _expressionBox.Focus(); _expressionBox.SelectionStart = _expressionBox.TextLength; };

            var hint = new Label
            {
                Text = UiText.ExprHint,
                Font = new Font("Microsoft YaHei UI", 9f),
                ForeColor = Color.FromArgb(97, 121, 153),
                AutoSize = true,
                Location = new Point(24, 68)
            };

            var syntax = new Label
            {
                Text = UiText.Syntax,
                Font = new Font("Microsoft YaHei UI", 10f),
                ForeColor = Color.FromArgb(77, 109, 152),
                AutoSize = true,
                Location = new Point(24, 94)
            };

            var quickPanel = new FlowLayoutPanel
            {
                Location = new Point(20, 128),
                Size = new Size(934, 80),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                WrapContents = true,
                AutoScroll = true
            };

            foreach (string token in new[]
            {
                "A", "B", "C", "D", "E", "F",
                "+", "-", "*", "/", "^2",
                "T()", "REF()", "RREF()", "RANK()", "TR()", "NORM()", "INV()", "DET()", "COF()", "ADJ()", "NULL()"
            })
            {
                quickPanel.Controls.Add(BuildQuickInsertButton(token));
            }

            Controls.Add(title);
            Controls.Add(_expressionBox);
            Controls.Add(runButton);
            Controls.Add(sampleButton);
            Controls.Add(hint);
            Controls.Add(syntax);
            Controls.Add(quickPanel);

            FormClosing += delegate(object sender, FormClosingEventArgs e)
            {
                if (!_allowClose && e.CloseReason == CloseReason.UserClosing)
                {
                    e.Cancel = true;
                    Hide();
                }
            };
        }

        public void FocusExpression()
        {
            if (!Visible)
            {
                Show();
            }

            BringToFront();
            Activate();
            _expressionBox.Focus();
            _expressionBox.SelectionStart = _expressionBox.TextLength;
        }

        public void AllowClose()
        {
            _allowClose = true;
        }

        private Button BuildQuickInsertButton(string token)
        {
            var button = new Button
            {
                Text = token,
                Size = new Size(token.Length > 4 ? 86 : 70, 34),
                FlatStyle = FlatStyle.Flat,
                BackColor = token.Contains("(") ? Color.FromArgb(232, 245, 244) : Color.FromArgb(236, 243, 250),
                ForeColor = Color.FromArgb(28, 76, 118),
                Font = new Font("Microsoft YaHei UI", 9f),
                Margin = new Padding(0, 0, 10, 10)
            };

            button.FlatAppearance.BorderSize = 0;
            button.Click += delegate
            {
                InsertToken(token);
            };
            return button;
        }

        private Button BuildActionButton(string text, Color backColor, Color foreColor, Point location, Size size)
        {
            var button = new Button
            {
                Text = text,
                Location = location,
                Size = size,
                FlatStyle = FlatStyle.Flat,
                BackColor = backColor,
                ForeColor = foreColor,
                Font = new Font("Microsoft YaHei UI", 10.5f, FontStyle.Bold)
            };

            button.FlatAppearance.BorderSize = 0;
            return button;
        }

        private void InsertToken(string token)
        {
            string text = token;
            int caretShift = 0;
            if (token.EndsWith("()", StringComparison.Ordinal))
            {
                text = token;
                caretShift = -1;
            }

            int start = _expressionBox.SelectionStart;
            _expressionBox.Text = _expressionBox.Text.Insert(start, text);
            _expressionBox.Focus();
            _expressionBox.SelectionStart = start + text.Length + caretShift;
        }
    }

    internal sealed class MainForm : Form
    {
        private readonly FlowLayoutPanel _matrixPanel;
        private Label _countLabel;
        private readonly Dictionary<string, TextBox> _matrixInputs;
        private readonly ResultForm _resultForm;
        private readonly ExpressionForm _expressionForm;
        private int _matrixCount;

        public MainForm()
        {
            _matrixInputs = new Dictionary<string, TextBox>();
            _matrixCount = 2;

            Text = UiText.Title;
            StartPosition = FormStartPosition.CenterScreen;
            Size = new Size(1260, 760);
            MinimumSize = new Size(980, 620);
            BackColor = Color.FromArgb(239, 244, 250);

            var header = new Panel
            {
                Dock = DockStyle.Top,
                Height = 124,
                BackColor = Color.FromArgb(35, 70, 116)
            };

            var title = new Label
            {
                Text = UiText.Title,
                Font = new Font("Microsoft YaHei UI", 30f, FontStyle.Bold),
                ForeColor = Color.White,
                AutoSize = true,
                Location = new Point(28, 18)
            };

            var subtitle = new Label
            {
                Text = UiText.Subtitle,
                Font = new Font("Microsoft YaHei UI", 13f),
                ForeColor = Color.FromArgb(228, 238, 251),
                AutoSize = true,
                Location = new Point(31, 74)
            };

            header.Controls.Add(title);
            header.Controls.Add(subtitle);

            _matrixPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 360,
                Padding = new Padding(24, 44, 24, 16),
                AutoScroll = true,
                WrapContents = true,
                FlowDirection = FlowDirection.LeftToRight,
                BackColor = Color.FromArgb(239, 244, 250)
            };

            var footer = BuildFooter();

            Controls.Add(footer);
            Controls.Add(_matrixPanel);
            Controls.Add(header);

            _resultForm = new ResultForm();
            _expressionForm = new ExpressionForm(this);

            CreateMatrixCards();

            Shown += delegate
            {
                _expressionForm.Location = new Point(Right + 10, Top + 70);
                _resultForm.Location = new Point(Right + 10, Top + 350);
                _expressionForm.Show(this);
                _resultForm.Show(this);
            };

            FormClosing += delegate
            {
                _expressionForm.AllowClose();
                _expressionForm.Close();
                _resultForm.Close();
            };
        }

        public void ExecuteExpression(string expression)
        {
            try
            {
                string expr = (expression ?? string.Empty).Trim();
                if (expr.Length == 0)
                {
                    return;
                }

                Dictionary<string, Matrix> context = BuildContext(expr);
                EvalValue value = new ExpressionParser(expr, context).Parse();
                _resultForm.ShowResult(UiText.Done, FormatValue(value));
            }
            catch (Exception ex)
            {
                _resultForm.ShowResult(UiText.EvalFailed, ex.Message);
            }
        }

        private Panel BuildFooter()
        {
            var footer = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(239, 244, 250),
                Padding = new Padding(20, 8, 20, 20)
            };

            var strip = new Panel
            {
                Height = 74,
                Dock = DockStyle.Top,
                BackColor = Color.White
            };

            var countText = new Label
            {
                Text = UiText.MatrixCount,
                Font = new Font("Microsoft YaHei UI", 17f, FontStyle.Bold),
                ForeColor = Color.Black,
                AutoSize = true,
                Location = new Point(18, 22)
            };

            _countLabel = new Label
            {
                Text = _matrixCount.ToString(CultureInfo.InvariantCulture),
                Font = new Font("Microsoft YaHei UI", 17f, FontStyle.Bold),
                ForeColor = Color.FromArgb(34, 94, 148),
                AutoSize = true,
                Location = new Point(142, 22)
            };

            var addButton = BuildFooterButton(UiText.AddMatrix, Color.FromArgb(41, 131, 156), Color.White, new Point(195, 16), new Size(92, 38));
            addButton.Click += delegate
            {
                if (_matrixCount >= 6)
                {
                    MessageBox.Show(UiText.MaxMatrices, UiText.Title, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                _matrixCount++;
                CreateMatrixCards();
            };

            var removeButton = BuildFooterButton(UiText.RemoveMatrix, Color.FromArgb(217, 231, 245), Color.FromArgb(51, 88, 126), new Point(300, 16), new Size(92, 38));
            removeButton.Click += delegate
            {
                if (_matrixCount > 1)
                {
                    _matrixCount--;
                    CreateMatrixCards();
                }
            };

            var exprButton = BuildFooterButton(UiText.OpenExpression, Color.FromArgb(236, 245, 244), Color.FromArgb(24, 106, 112), new Point(414, 16), new Size(126, 38));
            exprButton.Click += delegate { _expressionForm.FocusExpression(); };

            var resultButton = BuildFooterButton(UiText.OpenResult, Color.FromArgb(247, 240, 227), Color.FromArgb(135, 88, 14), new Point(553, 16), new Size(110, 38));
            resultButton.Click += delegate { _resultForm.Show(); _resultForm.BringToFront(); _resultForm.Activate(); };

            var syntax = new Label
            {
                Text = UiText.Syntax,
                Font = new Font("Microsoft YaHei UI", 10.5f),
                ForeColor = Color.FromArgb(77, 109, 152),
                AutoSize = true,
                Location = new Point(685, 24)
            };

            strip.Controls.Add(countText);
            strip.Controls.Add(_countLabel);
            strip.Controls.Add(addButton);
            strip.Controls.Add(removeButton);
            strip.Controls.Add(exprButton);
            strip.Controls.Add(resultButton);
            strip.Controls.Add(syntax);

            footer.Controls.Add(strip);
            return footer;
        }

        private Button BuildFooterButton(string text, Color backColor, Color foreColor, Point location, Size size)
        {
            var button = new Button
            {
                Text = text,
                Location = location,
                Size = size,
                FlatStyle = FlatStyle.Flat,
                BackColor = backColor,
                ForeColor = foreColor,
                Font = new Font("Microsoft YaHei UI", 9f)
            };

            button.FlatAppearance.BorderSize = 0;
            return button;
        }

        private void CreateMatrixCards()
        {
            var oldTexts = _matrixInputs.ToDictionary(pair => pair.Key, pair => pair.Value.Text);
            _matrixInputs.Clear();
            _matrixPanel.Controls.Clear();

            for (int i = 0; i < _matrixCount; i++)
            {
                string name = ((char)('A' + i)).ToString();
                var card = new Panel
                {
                    Size = new Size(290, 252),
                    BackColor = Color.White,
                    Margin = new Padding(0, 0, 18, 18)
                };

                var label = new Label
                {
                    Text = "Matrix " + name,
                    Font = new Font("Microsoft YaHei UI", 16f, FontStyle.Bold),
                    ForeColor = Color.Black,
                    AutoSize = true,
                    Location = new Point(18, 16)
                };

                var tip = new Label
                {
                    Text = UiText.InputTip,
                    Font = new Font("Microsoft YaHei UI", 9.5f),
                    ForeColor = Color.FromArgb(108, 132, 168),
                    AutoSize = true,
                    Location = new Point(18, 48)
                };

                var box = new TextBox
                {
                    Multiline = true,
                    ScrollBars = ScrollBars.Vertical,
                    Font = new Font("Consolas", 17f),
                    Location = new Point(18, 78),
                    Size = new Size(238, 156),
                    BorderStyle = BorderStyle.FixedSingle
                };

                string previous;
                if (oldTexts.TryGetValue(name, out previous))
                {
                    box.Text = previous;
                }
                else
                {
                    box.Text = i == 0 ? "[1 2; 3 4]" : i == 1 ? "[1 0; 0 1]" : string.Empty;
                }

                _matrixInputs[name] = box;
                card.Controls.Add(label);
                card.Controls.Add(tip);
                card.Controls.Add(box);
                _matrixPanel.Controls.Add(card);
            }

            _countLabel.Text = _matrixCount.ToString(CultureInfo.InvariantCulture);
        }

        private Dictionary<string, Matrix> BuildContext(string expression)
        {
            var names = Regex.Matches(expression ?? string.Empty, @"\b[A-Z]\b")
                .Cast<Match>()
                .Select(match => match.Value.ToUpperInvariant())
                .Distinct()
                .ToList();

            var context = new Dictionary<string, Matrix>(StringComparer.OrdinalIgnoreCase);
            foreach (string name in names)
            {
                TextBox box;
                if (!_matrixInputs.TryGetValue(name, out box))
                {
                    throw new InvalidOperationException(UiText.UnknownVar + name);
                }

                context[name] = MatrixInputParser.Parse(box.Text, "Matrix " + name);
            }

            return context;
        }

        private static string FormatValue(EvalValue value)
        {
            switch (value.Kind)
            {
                case ValueKind.Matrix:
                    return MatrixFormatter.ToMatlab(value.MatrixValue);
                case ValueKind.Scalar:
                    return NumberUtil.Format(value.ScalarValue);
                case ValueKind.Text:
                    return value.TextValue ?? string.Empty;
                default:
                    return string.Empty;
            }
        }
    }

    internal static class DpiUtil
    {
        [DllImport("user32.dll")]
        private static extern bool SetProcessDPIAware();

        public static void Enable()
        {
            try
            {
                SetProcessDPIAware();
            }
            catch
            {
            }
        }
    }

    internal static class Program
    {
        [STAThread]
        private static void Main()
        {
            DpiUtil.Enable();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}
