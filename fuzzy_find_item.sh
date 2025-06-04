#!/bin/bash

# Check if an argument is provided
if [ $# -ne 1 ]; then
    echo "Usage: $0 <name_to_search>"
    exit 1
fi

# Store the search term
search_term="$1"

# Search recursively for the item with fuzzy matching
# *search_term* will match anything that contains search_term
result=$(find . -name "*${search_term}*")

if [ -n "$result" ]; then
    found_count=0
    while IFS= read -r item; do
        if [ -f "$item" ]; then
            echo "found, the requested is file"
            echo "name: $(basename "$item")"
            echo "path: $(realpath "$item")"
            echo "---"
            found_count=$((found_count + 1))
        elif [ -d "$item" ]; then
            echo "found, the requested is dir"
            echo "name: $(basename "$item")"
            echo "path: $(realpath "$item")"
            echo "---"
            found_count=$((found_count + 1))
        fi
    done <<< "$result"
    echo "Total matches found: $found_count"
else
    echo "not found"
fi 