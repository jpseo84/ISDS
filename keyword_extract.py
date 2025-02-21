#!/usr/bin/env python3
"""
Keyword Extractor for Cybersecurity & ISDS Regulation

Copyright (C) 2024 Jupyo Seo

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <https://www.gnu.org/licenses/>.

For inquiries, contact: github.com/jpseo84/ISDS
"""

import csv
import time
import argparse
import os
import sys

def load_keywords(keyword_file):

    it_keywords = set()
    is_keywords = set()

    try:
        with open(keyword_file, "r", encoding="utf-8") as csvfile:
            reader = csv.reader(csvfile)
            next(reader)

            for row in reader:
                if not row or row[0].startswith("#"):
                    continue
                if len(row) < 3:
                    continue
                keyword, category, _ = row 
                if category == "IT":
                    it_keywords.add(keyword.lower())
                elif category == "IS":
                    is_keywords.add(keyword.lower())

        return it_keywords, is_keywords

    except FileNotFoundError:
        print(f"Error: Missing keyword file - '{keyword_file}' not found.")
        sys.exit(1)

def process_large_file(target_file, keyword_file):

    it_keywords, is_keywords = load_keywords(keyword_file)

    it_output = "IT_Output.csv"
    is_output = "IS_Output.csv"

    start_time = time.time()
    total_lines, it_lines, is_lines = 0, 0, 0

    try:
        with open(target_file, "r", encoding="utf-8") as infile, \
             open(it_output, "w", encoding="utf-8", newline="") as it_out, \
             open(is_output, "w", encoding="utf-8", newline="") as is_out:
            
            it_writer = csv.writer(it_out)
            is_writer = csv.writer(is_out)

            for line in infile:
                total_lines += 1
                words = set(line.lower().split())

                it_match = any(word in it_keywords for word in words)
                is_match = any(word in is_keywords for word in words)

                if it_match:
                    it_writer.writerow([line.strip()])
                    it_lines += 1

                if is_match:
                    is_writer.writerow([line.strip()])
                    is_lines += 1

                if total_lines % 1_000_000 == 0:
                    elapsed_time = time.time() - start_time
                    print(f"Processed {total_lines:,} lines... Time elapsed: {elapsed_time:.2f} sec")

        elapsed_time = time.time() - start_time
        print("\n File Processing complete!")
        print(f"Total lines processed: {total_lines:,}")
        print(f"IT-related lines written: {it_lines:,}")
        print(f"IS-related lines written: {is_lines:,}")
        print(f"Total processing time: {elapsed_time:.2f} sec")

    except FileNotFoundError:
        print(f"Error: Target file '{target_file}' not found.")
        sys.exit(1)

def main():
    print("Keyword Extractor for ISDS")
    print("Copyright (C) 2024 Jupyo Seo")
    print("This program comes with ABSOLUTELY NO WARRANTY.")
    print("This is free software, and you are welcome to redistribute it")
    print("under certain conditions. See <https://www.gnu.org/licenses/> for details.\n")
    
    parser = argparse.ArgumentParser(description="Keyword-based text file extractor.")
    parser.add_argument("-t", "--target", type=str, help="Target text file to process")
    parser.add_argument("-k", "--keyword", type=str, help="CSV file containing keyword data")
    parser.add_argument("positional_target", nargs="?", type=str, help="Alternative way to specify the target file")
    args = parser.parse_args()

    target_file = args.target or args.positional_target
    if not target_file:
        print("Error: No target file specified.")
        print("Usage:")
        print("  ./keyword_extract.py -t <targetfile> -k <keywordfile>")
        print("  ./keyword_extract.py --target <targetfile> --keyword <keywordfile>")
        print("  ./keyword_extract.py <targetfile>  (When a 'keywords.csv' is the same directory)")
        sys.exit(1)

    keyword_file = args.keyword if args.keyword else "keywords.csv"

    if not os.path.exists(keyword_file):
        print(f"Error: Missing keyword file - '{keyword_file}' not found.")
        sys.exit(1)

    process_large_file(target_file, keyword_file)

if __name__ == "__main__":
    main()