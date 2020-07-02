# SlideValidator

## Purpose
The SlideValidator will provide automation support for validating large presentations. A set of rules will be applied to each slide and if the rule isn't matched then a comment will be added to the slide. Every rule has a matching configuration slide in the SlideValidator presentation.

Currently there is just on rule for detecting the usage of non-permitted fonts.

## Background
Beside the obvious purpose SlideValidator also introduces en example driven test-framework supporting a very small subset of [Gherkin](https://cucumber.io/docs/gherkin/reference/). In the future I will move the test-framework into it's own repository.

## Usage
Please be aware that Microsoft Office macros are considered as a security risk. Even [Microsoft says so](https://support.microsoft.com/en-us/office/enable-or-disable-macros-in-office-files-12b036fd-d140-4e74-b45e-16fed1a7e5c6). Therefore

USE THIS APPLICATION AT YOUR OWN RISK!

To validate a presentation, download SlideValidator.pptm and run the macro "validate_presentation".

## Adapting SlideValidator
Of course you may clone SlideValidator and add your own rules.

More details coming soon... 
