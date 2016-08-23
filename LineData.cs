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
namespace ocli
{
    public class LineData
    {
        public static int LongestFrom = 0;

        public string Id { get; private set; }
        public string Age { get; private set; }
        public string From { get; private set; }
        public string Title { get; private set; }

        public LineData(string id, string age, string from, string title)
        {
            Id = id;
            Age = age;
            From = from;
            Title = title;

            LongestFrom = (LongestFrom < From.Length ? From.Length : LongestFrom);
        }
    }
}
