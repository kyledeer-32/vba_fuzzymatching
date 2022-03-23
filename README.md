# **vba_fuzzymatching**
=========================
## **Introduction:**
-------------------------
These VBA functions enable Excel users to execute fuzzy matching by using the basic algorithm for computing the **Levenshtein Distance** between two strings. Additionally, string similarity can be calculated between strings as a method of determining how strong matches are.

## **Functions**
-------------------------
### =LevD([String 1],[String 2])
Calculates the Levenshtein Distance between two strings. The Levenshtein Distance is the minimum number of character insertions, deletions, and substitutions required to perfectly match one string to another. This formula isn't case sensitive.

### =Fuzzy_Match([String 1],[Array of Strings]) 
Traverses an array of strings, calculating the Levenshtein Distance between each and [String 1] (String 1 is the string you are seeking the closest match for in the [Array of Strings]. Returns the closest match, i.e., the string with the lowest Levenshtein Distance. If two or more strings have lowest Levenshtein distance, the first one traversed in the [Array of Strings] will be returned.

### =String_Similarity([String 1],[String 2])
Calculates the similarity percentage between two strings using the Levenshtein Distance, e.g., "rock" and "sock" are 75% similiar, because 3 of their 4 characters are the same.

## **Limitations/In-Development**
-------------------------
1. There is no explicit parameter for setting the matching threshold, i.e., the "Fuzzy_Match" function will return the closest match, even if it's poor.
2. There is no mechanism for partitioning phrases, sentences, etc. Currently, one phrase (more than one word) is treated as one string.

## **Technologies Used**
-------------------------
Visual Basic Editor for Excel

## **References**
-------------------------
Levenshtein Distance Formula: https://en.wikipedia.org/wiki/Levenshtein_distance?msclkid=154f886daafe11ec82b22174e1344ccf
Basic string similarity formula: http://adamfortuno.com/index.php/2021/07/05/levenshtein-distance-and-distance-similarity-functions/?msclkid=aa598709a8b611ec8de2ac84df81a9da

