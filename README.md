# PowerPoint VBA Name Card Generator

A VBA automation project that generates personalized name cards in Microsoft PowerPoint from a reusable slide template and a list of names.

> **Generate the logic once. Execute the repetition locally.**

[Versión en español](README.es.md)

## Overview

This project automates the creation of personalized name cards in Microsoft PowerPoint.

The macro uses a single slide as a template, generates one copy for each name in a predefined list, and modifies only the text box named `NOMBRE_INTEGRANTE`.

All other visual elements, including logos, lines, formatting, positioning, and slide dimensions, remain unchanged.

## The Problem

Creating personalized name cards manually requires duplicating a slide, replacing the name, and verifying the formatting of every copy.

When the list includes many participants, this process becomes repetitive and increases the risk of:

- Omitting names.
- Duplicating records.
- Accidentally modifying the design.
- Producing inconsistent formatting.
- Spending operational time on a low-value repetitive task.

## The Solution

The automation combines:

- A reusable PowerPoint slide template.
- A text box identified as `NOMBRE_INTEGRANTE`.
- A delimiter-separated list of names.
- A VBA macro that duplicates the template and replaces the text.
- Automatic text resizing to keep each name on a single line.
- Basic validations to prevent incorrect execution.

The result is one automatically generated slide for each name in the list.

## Demo

### Template

The presentation begins with a single template slide:

![PowerPoint template](ASSETS/plantilla-demo.jpg)

### Generated Result

After running the macro, PowerPoint automatically generates one slide for each name:

![Automatically generated slides](ASSETS/resultado-generado.jpg)

## Features

- Automatically generates one slide per name.
- Preserves the original visual design.
- Validates that the presentation contains only one template slide.
- Validates the existence of the `NOMBRE_INTEGRANTE` object.
- Keeps each name on a single line.
- Automatically reduces the font size when necessary.
- Supports accented characters and multi-word names.
- Calculates the number of generated slides dynamically.
- Displays a confirmation message after execution.
- Runs locally inside Microsoft PowerPoint.

## Project Structure

```text
powerpoint-vba-name-card-generator/
├── README.md
├── README.es.md
├── ASSETS/
│   ├── plantilla-demo.jpg
│   └── resultado-generado.jpg
├── DEMO/
│   └── Plantilla_Generador_Letreros_Demo.pptm
└── SRC/
    └── GeneradorLetreros.bas
