# **vba_fuzzymatching**

## **Introduction:**
-------------------------
This VBA module contains 4 user-defined functions (UDFs) that enable Excel users to execute fuzzy matching by using the basic algorithm for computing the **Levenshtein Distance** between two strings. Following are details on these functions.

## **Functions**
-------------------------
### =LevD([String 1],[String 2])
Calculates the Levenshtein Distance between two strings. The Levenshtein Distance is the minimum number of character insertions, deletions, and substitutions required to perfectly match one string to another. This formula isn't case sensitive.

### =Fuzzy_Match([String 1],[Array of Strings], [Threshold]) 
Traverses an array of strings, calculating the Levenshtein Distance between each item (string) in the [Array of Strings] and [String 1] (String 1 is the string you are seeking a match for). If a value in the array of strings is a sentence (includes a space), then it will be partitioned and matched based on the best-matching substring (see "bestword" function below). The third paramter: [Threshold], requires an input value between 0 and 1. This is the minimum string similiarity the user desires for fuzzy matching, e.g., if the user inputs ".75", then no matches will be returned with a string similarity less than 75%. This function Returns the closest match, i.e., the string with the lowest Levenshtein Distance.

Note: If two or more strings have lowest Levenshtein distance, the first one traversed in the [Array of Strings] will be returned.

### =bestword([string1], [string2])
Takes a sentence ([string2] and partitions it using delimeter = (" "). Then each substring will be matched to [String1]. The best matching substring will be returned.

### =String_Similarity([String 1],[String 2])
Calculates the similarity percentage between two strings using the Levenshtein Distance, e.g., "rock" and "sock" are 75% similiar, because 3 of their 4 characters are the same.

## **Technologies Used**
-------------------------
Visual Basic Editor for Excel

## **References**
-------------------------
Levenshtein Distance Formula: https://en.wikipedia.org/wiki/Levenshtein_distance?msclkid=154f886daafe11ec82b22174e1344ccf

Basic string similarity formula: http://adamfortuno.com/index.php/2021/07/05/levenshtein-distance-and-distance-similarity-functions/?msclkid=aa598709a8b611ec8de2ac84df81a9da

