# VBA-Island-Challenge
VBA case study challenge to desing a method for counting "Islands" on a binary grid.

## Instructions
Given a 2-d grid map of '1's (land) and '0's (water), count the number of islands. 

An island is surrounded by water and is formed by connecting adjacent lands horizontally or vertically. 

You may assume all four edges of the grid are all surrounded by water.

| A     | B     | C     | D     | E     | F     | G     | H     | I     | J     | K     | 
| :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | 
| 0     | 0     | 0     | 0     | 0     | 0     | 0     | 0     | 0     | 0     | 0     | 
| 0     | 0     | 0     | 1     | 1     | 0     | 1     | 0     | 1     | 0     | 0     | 
| 0     | 1     | 1     | 1     | 0     | 0     | 1     | 1     | 1     | 1     | 0     | 
| 0     | 1     | 0     | 0     | 1     | 1     | 0     | 0     | 0     | 1     | 0     | 
| 0     | 1     | 1     | 1     | 1     | 0     | 1     | 1     | 1     | 0     | 0     | 
| 0     | 1     | 1     | 1     | 1     | 0     | 0     | 1     | 0     | 1     | 0     | 
| 0     | 1     | 0     | 1     | 0     | 1     | 0     | 0     | 0     | 0     | 0     | 
| 0     | 1     | 0     | 0     | 1     | 1     | 0     | 0     | 1     | 1     | 0     | 
| 0     | 1     | 1     | 1     | 0     | 0     | 1     | 1     | 1     | 1     | 0     | 
| 0     | 1     | 0     | 0     | 0     | 1     | 1     | 1     | 1     | 1     | 0     | 
| 0     | 0     | 1     | 1     | 0     | 0     | 0     | 0     | 1     | 1     | 0     | 
| 0     | 1     | 1     | 1     | 1     | 1     | 0     | 1     | 0     | 0     | 0     | 
| 0     | 0     | 1     | 1     | 0     | 1     | 0     | 0     | 1     | 1     | 0     | 
| 0     | 0     | 1     | 1     | 0     | 1     | 1     | 1     | 0     | 1     | 0     | 
| 0     | 1     | 1     | 0     | 1     | 1     | 1     | 0     | 1     | 0     | 0     | 
| 0     | 1     | 1     | 1     | 0     | 0     | 1     | 1     | 1     | 0     | 0     | 
| 0     | 1     | 1     | 0     | 1     | 0     | 1     | 1     | 1     | 0     | 0     | 
| 0     | 1     | 1     | 1     | 0     | 0     | 1     | 0     | 0     | 1     | 0     | 
| 0     | 0     | 0     | 0     | 0     | 0     | 0     | 0     | 0     | 0     | 0     | 
