## requirements

- python 3.6+
- anthropic API key
- powerPoint file in `.pptx` format

## installation

1. install required packages:
```bash
pip install python-pptx requests
```

2. paste anthropic api key in the simple_keyword_extractor.py file

## usage

1. **configure the script:**
   - open `keyword_extractor.py`
   - replace `""` with your actual Anthropic API key
   - Update `PPT_FILE` path to your PowerPoint file

2. **run the script:**
```bash
python keyword_extractor.py
```

3. **check the output:**
   - the script creates a new file: `your_file_highlighted.pptx`
   - important keywords are highlighted in red and green, can be changed based on preference
   - console shows progress and extracted keywords

## example Output

```
Processing: input/presentation.pptx
Found 3 slides
Slide 1: Getting keywords...
Slide 1: Keywords=['interpreted', 'interactive', 'object-oriented', 'readable', 'syntactical']
Slide 1: Applied 5 highlights
Slide 2: Keywords=['runtime', 'interpreter', 'compile', 'blocks', 'statements']
Slide 2: Applied 4 highlights
Complete! Total highlights: 9
Saved to: input/presentation_highlighted.pptx
```

## configuration

you can modify these settings in the script:

- `MAX_KEYWORDS`: humber of keywords per slide (default: 5)
- `RED/GREEN`: Highlight colors
- file paths and API settings

## how It Works

1. extracts text from each slide in your PowerPoint
2. sends slide content to AI with specific instructions to find study-relevant keywords
3. validates that keywords actually exist in the slide content
4. highlights unique keywords with alternating red/green colors
5. saves the highlighted presentation as a new file

## troubleshooting

**"API error"**: check your API key and internet connection

**"no keywords found"**: slide might have no text or API couldn't find relevant terms

**"file not found"**: verify the PowerPoint file path is correct

**"import error"**: run `pip install python-pptx requests`
