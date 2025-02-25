# PowerPoint Skeleton Generator

## Overview
This script automates the creation of a PowerPoint presentation skeleton from a JSON input. It helps users quickly generate structured slides with titles, key points, and optional images. The generated slides serve as a foundation that users can refine and enhance later.

## Features
- Creates a PowerPoint presentation from a JSON file.
- Supports images (768x960 recommended) if they exist; skips them if not.
- Titles are auto-centered, and content is aligned based on JSON specifications.
- Supports a background image for styling.
- Ensures a consistent, styled layout using predefined font settings.

## Installation
To set up the environment and install dependencies:
```sh
python -m venv pptmaker
cd pptmaker
scripts\activate  # Use 'source scripts/activate' on macOS/Linux
python -m pip install --upgrade pip
python -m pip install python-pptx
```

## Usage
To run the script:
```sh
python pptmaker.py
```

Make sure to provide a valid JSON file named `slides_data.json` in the script directory.

## JSON Format
The script expects a structured JSON file like this:
```json
{
    "Title Pg 1": {
      "alignment": "right",
      "image": "image.png",
      "content": {
        "LLM Usage": "Your Skill Up",
        "Security Awareness": "Summaries Only",
        "Application Security": "Owasp Top 10 GenAI",
        "ML Processes": "Understand Weak Points",
        "Overall Goal": "Encourage Policy Implementation"
      }
    },
    "Prompting Basics": {
      "alignment": "left",
      "image": "prompting_basics.png",
      "content": {
        "Limit Scope of Questions": "Start new Chats, Ask Specific questions, Understand objective",
        "Prompt Structuring": "Question First, Specifics Next, Examples, Whitelist, Suggest Steps",
        "Contextual Guidance": "Examples & Delimiters/XML/Mark Down, Desired Output",
        "Iterative Refinement": "Guide, Refine, Challenge, What was missed?"
      }
    }
}
```

## Limitations
- **Unique Titles**: All slide titles must be unique; otherwise, the JSON will not be valid.
- **Image Placement**: Images should be **768x960** for optimal placement, otherwise, they may not fit well.
- **Content Density**: It is recommended to limit **each slide to 3-4 key points** to avoid overcrowding.
- **Error Handling**: The script skips missing images and prints a warning but continues processing.

## Enhancements and Refinements
This tool provides a structured starting point, but manual refinement is encouraged to improve slide aesthetics, adjust layouts, and enhance readability.

## License
This script is provided as-is with no warranties. Modify and use it at your discretion.

---
**Author:** Generated for automation and efficiency in creating PowerPoint slides.
