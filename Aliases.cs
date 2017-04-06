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
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ocli
{
    public class Aliases
    {
        public Dictionary<string, string> NameAlias { get; private set; }

        public Aliases(string file)
        {
            NameAlias = new Dictionary<string, string>();

            var query = File.ReadAllLines(file)
                            .Where(l => l.Length > 1 && l.Substring(0, 1) != "#");

            var allAliases = query.ToList();

            try
            {
                foreach (var line in allAliases)
                {
                    string[] columns = line.Split(':');
                    NameAlias.Add(columns[0].Trim(), columns[1].Trim());
                }
            }
            catch (System.Exception)
            {
                //If anything goes wrong set the alias dictionary to empty
                NameAlias.Clear();
            }
        }
    }
}
