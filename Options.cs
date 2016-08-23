/*  Copyright (c) 2016 xanthalas.co.uk
 * 
 *  Author: Xanthalas
 *  Date  : August 2016
 * 
 *  This file is part of ocli
 *
 *  ocli is free software: you can redistribute it and/or modify
 *  it under the terms of the GNU General Public License as published by
 *  the Free Software Foundation, either version 3 of the License, or
 *  (at your option) any later version.
 *
 *  ocli is distributed in the hope that it will be useful,
 *  but WITHOUT ANY WARRANTY; without even the implied warranty of
 *  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 *  GNU General Public License for more details.
 *
 *  You should have received a copy of the GNU General Public License
 *  along with ocli.  If not, see <http://www.gnu.org/licenses/>.
 */
using CommandLine;

namespace ocli
{
    public class Options
    {
        [Option('h', "help", Required = false, HelpText = "Show this help.")]
        public bool Help { get; set; }

        [Option('a', "all", DefaultValue = false, HelpText = "Show all emails")]
        public bool ShowAll { get; set; }

        [Option('t', "today", DefaultValue = false, HelpText = "Show all emails received today")]
        public bool Today { get; set; }
    }
}
