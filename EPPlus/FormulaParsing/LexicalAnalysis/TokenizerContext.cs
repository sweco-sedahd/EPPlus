﻿/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See https://github.com/JanKallman/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied.
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 *
 * Author							Change						Date
 * ******************************************************************************
 * Mats Alm   		                Added       		        2013-03-01 (Prior file history on https://github.com/swmal/ExcelFormulaParser)
 *******************************************************************************/

using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis
{
    public class TokenizerContext
    {
        private StringBuilder _currentToken;
        private readonly List<Token> _result;

        public TokenizerContext(string formula)
        {
            if (!string.IsNullOrEmpty(formula))
            {
                FormulaChars = formula.ToArray();
            }

            _result = new List<Token>();
            _currentToken = new StringBuilder();
        }

        public char[] FormulaChars { get; }

        public IList<Token> Result => _result;

        public bool IsInString { get; private set; }

        public bool IsInSheetName { get; private set; }

        internal int BracketCount { get; set; }

        public string CurrentToken => _currentToken.ToString();

        public bool CurrentTokenHasValue => !string.IsNullOrEmpty(IsInString ? CurrentToken : CurrentToken.Trim());

        public Token LastToken => _result.Count > 0 ? _result.Last() : null;

        public void ToggleIsInString()
        {
            IsInString = !IsInString;
        }

        public void ToggleIsInSheetName()
        {
            IsInSheetName = !IsInSheetName;
        }

        public void NewToken()
        {
            _currentToken = new StringBuilder();
        }

        public void AddToken(Token token)
        {
            _result.Add(token);
        }

        public void AppendToCurrentToken(char c)
        {
            _currentToken.Append(c.ToString());
        }

        public void AppendToLastToken(string stringToAppend)
        {
            _result.Last().Append(stringToAppend);
        }

        public void SetLastTokenType(TokenType type)
        {
            _result.Last().TokenType = type;
        }

        public void ReplaceLastToken(Token newToken)
        {
            int count = _result.Count;
            if (count > 0)
            {
                _result.RemoveAt(count - 1);
            }

            _result.Add(newToken);
        }
    }
}