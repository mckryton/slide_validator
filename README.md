# SlideValidator

## Purpose
The SlideValidator will provide automation support for validating large presentations. A set of rules will be applied to each slide and if the rule isn't matched then a comment will be added to the slide. Every rule has a matching configuration slide in the SlideValidator presentation.

Currently there is just on rule for detecting the usage of non-permitted fonts.

## Background
Beside the obvious purpose SlideValidator also introduces en example driven test-framework supporting a very small subset of [Gherkin](https://cucumber.io/docs/gherkin/reference/). In the future I will move the test-framework into it's own repository.

## Security
Unfortunately macros for Microsoft Office are considered as a security risk. Even [Microsoft says so](https://support.microsoft.com/en-us/office/enable-or-disable-macros-in-office-files-12b036fd-d140-4e74-b45e-16fed1a7e5c6).
Because of their flexibility they are often used to as malware. So if Powerpoint warns you that your are activating an macro, please ensure that it's really your intention to do this.  

So be aware **USE THIS APPLICATION AT YOUR OWN RISK!**

## Usage
To validate a presentation, download SlideValidator.pptm and run the macro "validate_presentation".

## Adapting SlideValidator
Of course you may clone SlideValidator and add your own rules. <br> *Hint: This requires knowledge about how to write code in VBA.*

### Start with an example
#### Add a new feature
To describe your new rule copy the class [TTemplate](blob/master/source/TTemplate.cls) and name it Feature_Rule_<*your_rule_name*>. Open the function ```test_examples``` and fill in the examples describing the functionality of your new rule. In the declaration section add the tag "wip" (for work in progress) like this: ```Const cTags = "wip"```
<br>Hint: remove the wip tag from all other features. I might have forgotten to do this. ;-)

#### Register the new feature
Now open the procedure ```run_acceptance_tests```from the module [Teststart](blob/master/source/Teststart.bas). Add your feature to the list of test cases by modifying this line:<br>```acceptance_testcases = Array(New Feature_Rule_Permitted_Fonts, New Feature_Rule_<your_rule_name>)```
<br> Open the Immediate Window and enter ```Teststart.run_acceptance_tests "wip"```
<br> You should see something like <br> MISSING *your first step from the example* <br> Perfect! You are ready to implement the new rule.

### Create a rule
#### Add a new rule class
At the  moment SlideValidator does not provide a template for rules. So copy the class [Rule_Permitted_Fonts](blob/master/source/Rule_Permitted_Fonts.cls) and name after your new rule. You might already delete the ```collect_permitted_fonts``` function as well as the ```permitted_fonts``` property because both are unique to the permitted_fonts rule.

#### Register the rule
To make SlideValidator aware about the new rule add this line <br> ```Public <your_rule_name> As New Rule_<your_rule_name>``` to the [RuleCatalog](blob/master/source/RuleCatalog.cls) class.

#### Add a config slide
SlideValidator applies only those rule with a matching config slide. <br>
*Hint: Hidden config slides are ignored.* <br>
So just copy the config slide for the permitted fonts rule and adapt the title to match the name of your rule. While validating slides SlideValidator will read all the parameters from table in the config slide and make them available as strings in the config property of your rule class.

### Implement the rule logic
#### Make the steps from your examples pass the test
Finally you are ready to make your rule work! Open the Feature class you have added and go to the ```run_step``` procedure. Add a new case step matching the first step from your example. Look at the same method from the [Feature_Rule_Permitted_Fonts](blob/master/source/Feature_Rule_Permitted_Fonts.cls) class to get an idea how to fill the steps. <br> *Hint: have a look at the [TSupport](blob/master/source/TSupport.bas) module to find some useful helper function for writing tests.*

That's it! After making all your example steps pass your new rule is ready to be applied against your presentations.
