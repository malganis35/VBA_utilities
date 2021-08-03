# [How to Create a Progress Bar in Microsoft PowerPoint](https://www.howtogeek.com/709523/how-to-create-a-progress-bar-in-microsoft-powerpoint/)



Source : https://www.howtogeek.com/709523/how-to-create-a-progress-bar-in-microsoft-powerpoint/



![img](https://www.howtogeek.com/thumbcache/60/60/8e1b09679e2bf3e3afc95881195769c3/wp-content/uploads/2021/06/Marshall-Gunnell.png)[Marshall Gunnell](https://www.howtogeek.com/author/marshallgunnell/)

![img](https://www.howtogeek.com/wp-content/uploads/2021/06/Marshall-Gunnell.png)









 [@Marshall_G08](https://twitter.com/Marshall_G08)

​	Feb 23, 2021, 10:24 am EDT | 1 min read

![Microsoft PowerPoint Logo](https://www.howtogeek.com/wp-content/uploads/2019/07/stock-lede-microsoft-office-powerpoint-3.png?width=1198&trim=1,1&bg-color=000&pad=1,1)

A progress bar is a graphic that, in PowerPoint, visually represents  the percentage of the slideshow that has been completed. It’s also a  good indicator of the remaining amount. Here’s how to create a progress  bar in Microsoft PowerPoint.

You can manually create a progress bar by [inserting a shape](https://www.howtogeek.com/439038/how-to-insert-a-picture-or-other-object-in-microsoft-office/) at the bottom of each slide. The problem with this approach is that  you’ll need to measure the length of each shape based on the number of  slides in the presentation. Additionally, if you add or remove a slide,  you’ll need to manually redo the progress bar on every slide in the  slideshow.

<iframe id="google_ads_iframe_/10518929/tmnp.howtogeek/article_details/a0-p1-s2_1" title="3rd party ad content" name="google_ads_iframe_/10518929/tmnp.howtogeek/article_details/a0-p1-s2_1" scrolling="no" marginwidth="0" marginheight="0" style="border: 0px none; vertical-align: bottom;" sandbox="allow-forms allow-popups allow-popups-to-escape-sandbox allow-same-origin allow-scripts allow-top-navigation-by-user-activation" srcdoc="" data-google-container-id="7" data-load-complete="true" width="300" height="250" frameborder="0"></iframe>

To keep everything consistent and save yourself a serious headache, you can [use a macro](https://www.howtogeek.com/706392/how-to-enable-and-disable-macros-in-microsoft-office-365/) to create a progress bar. With this macro, the progress bar will  automatically adjust itself based on the number of slides in the  presentation.

**RELATED:** [***Macros Explained: Why Microsoft Office Files Can Be Dangerous\***](https://www.howtogeek.com/171993/macros-explained-why-microsoft-office-files-can-be-dangerous/)

First, [open the PowerPoint presentation](https://www.howtogeek.com/393248/what-is-a-pptx-file-and-how-do-i-open-one/) that you would like to create a progress bar for. Once it’s open, click the “View” tab, then select “Macros.”

![Macros option](https://www.howtogeek.com/wp-content/uploads/2021/01/Macros-option.png?trim=1,1&bg-color=000&pad=1,1)

The “Macro” window will appear. In the text box below “Macro Name,”  type in a name for your new macro. The name can’t contain spaces. Once  it’s entered, click “Create,” or, if you’re using Mac, click the “+”  icon.

![Enter macro name and click create](https://www.howtogeek.com/wp-content/uploads/2021/01/Enter-macro-name-and-click-create.png?trim=1,1&bg-color=000&pad=1,1)

Advertisement

<iframe id="google_ads_iframe_/10518929/tmnp.howtogeek/article_details/a1-p3-s2_0" title="3rd party ad content" name="google_ads_iframe_/10518929/tmnp.howtogeek/article_details/a1-p3-s2_0" scrolling="no" marginwidth="0" marginheight="0" style="border: 0px none; vertical-align: bottom;" sandbox="allow-forms allow-popups allow-popups-to-escape-sandbox allow-same-origin allow-scripts allow-top-navigation-by-user-activation" srcdoc="" data-google-container-id="8" data-load-complete="true" width="300" height="250" frameborder="0"></iframe>

The “Microsoft Visual Basic for Applications (VBA)” window will now open. In the editor, you’ll see this code:

```
Sub ProgressBar()

End Sub
```

![Macro code in editor](https://www.howtogeek.com/wp-content/uploads/2021/01/Macro-code-in-editor.png?trim=1,1&bg-color=000&pad=1,1)

First, place your cursor between the two lines of code.

![Cursor placement in editor](https://www.howtogeek.com/wp-content/uploads/2021/01/Cursor-placement-in-editor.png?trim=1,1&bg-color=000&pad=1,1)

Next, copy and paste this code:

```
On Error Resume Next
With ActivePresentation
For X = 1 To .Slides.Count
.Slides(X).Shapes("PB").Delete
Set s = .Slides(X).Shapes.AddShape(msoShapeRectangle, _
0, .PageSetup.SlideHeight - 12, _
X * .PageSetup.SlideWidth / .Slides.Count, 12)
s.Fill.ForeColor.RGB = RGB(127, 0, 0)
s.Name = "PB"
Next X:
End With
```

Once it’s pasted, your code should look like this in the editor.

![Final code format in the editor](https://www.howtogeek.com/wp-content/uploads/2021/01/Final-code-format-in-the-editor.png?trim=1,1&bg-color=000&pad=1,1)

> **Note:** There are no line breaks now between the first and last line of code.

You can now close the VBA window. Back in Microsoft PowerPoint, click “Macros” in the “View” tab again.

The Best Tech Newsletter Anywhere

Join **425,000** subscribers and get a daily digest of features, articles, news, and trivia.

​                                                

By submitting your email, you agree to the [Terms of Use](https://www.howtogeek.com/terms-of-use) and [Privacy Policy](https://www.howtogeek.com/privacy-policy).

![Macros option](https://www.howtogeek.com/wp-content/uploads/2021/01/Macros-option.png?trim=1,1&bg-color=000&pad=1,1)

Next, choose your macro name (“ProgressBar” in our example) to select it, then click “Run.”

![Select ProgressBar macro](https://www.howtogeek.com/wp-content/uploads/2021/01/Select-ProgressBar-macro.png?trim=1,1&bg-color=000&pad=1,1)

The progress bar will now appear at the bottom of each slide of your presentation.



Advertisement

<iframe id="google_ads_iframe_/10518929/tmnp.howtogeek/article_details/a1-p4-s2_0" title="3rd party ad content" name="google_ads_iframe_/10518929/tmnp.howtogeek/article_details/a1-p4-s2_0" scrolling="no" marginwidth="0" marginheight="0" style="border: 0px none; vertical-align: bottom;" sandbox="allow-forms allow-popups allow-popups-to-escape-sandbox allow-same-origin allow-scripts allow-top-navigation-by-user-activation" srcdoc="" data-google-container-id="9" data-load-complete="true" width="300" height="250" frameborder="0"></iframe>

If you delete a slide, the progress bar will adjust itself  automatically. If you add a new slide, you’ll need to run the macro  again (View > Macro > Run). It’s a minor inconvenience when  compared to adjusting everything manually.