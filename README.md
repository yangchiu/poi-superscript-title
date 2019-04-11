# Tag Superscripts and Titles in a Journal Paper doc/docx File with POI
- Based on Apache POI

- Check the vertical alignment enum of each word to determine it's a superscript or not.
- If the font size of a word is larger than the average size, it would be regarded as a title. 

- If a superscript found in the document, it would be tagged as \<sup>superscript\</sup>
- If a title found in the document, it would be tagged as <title>title</title> 

#### Dependency
* Apache POI

#### Usage
Java:
```python
$java -jar poi_sup_title.jar your_word_file
```
Python:
```python
$python poi_sup_title.py your_word_file
```
