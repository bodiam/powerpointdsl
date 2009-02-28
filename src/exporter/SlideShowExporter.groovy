package exporter

import builder.ImageSlide
import builder.Slide
import builder.SlideShow
import java.awt.Rectangle
import org.apache.poi.hslf.model.Picture
import org.apache.poi.hslf.model.Slide
import org.apache.poi.hslf.model.TextBox
import org.apache.poi.hslf.usermodel.RichTextRun
import org.apache.poi.hslf.usermodel.SlideShow

/**
 * @author Erik Pragt
 */
public class SlideShowExporter {
    def export(builder.SlideShow source) {
        SlideShow slideShow = new SlideShow()

        source.slides.each {sourceSlide ->
            // Create slide
            if (sourceSlide instanceof builder.Slide) {
                def slide = createSlide(sourceSlide.title, slideShow)

                // Create bullets
                if (sourceSlide.bullets) {
                    def text = sourceSlide.bullets.text.join('\n')
                    def bullets = createBullets(text)

                    // Add content to slide
                    slide.addShape(bullets)
                }

                // Create text
                if (sourceSlide.text) {
                    def textBox = createTextbox(sourceSlide.text)
                    slide.addShape(textBox)
                }
            } else if (sourceSlide instanceof builder.ImageSlide) {
                def slide = createImageSlide("", sourceSlide.src, slideShow)
            }
        }

        writeFile(slideShow, source.fileName)
    }


    /**
     * Creates a slide with a title
     */
    private def createSlide(title, SlideShow slideShow) {
        def slide = slideShow.createSlide()
        TextBox textBox = slide.addTitle()
        textBox.text = title
        return slide
    }

    private def createImageSlide(title, source, SlideShow slideShow) {
        int idx =  slideShow.addPicture(new File(new String(source)), Picture.JPEG)

        Picture pict = new Picture(idx);

        //set image position in the slide
        pict.setAnchor(new java.awt.Rectangle(0, 0, 800, 600));

        Slide slide = slideShow.createSlide();
        slide.addShape(pict);

    }

    /**
     * Creates a shape with bullets
     */
    private def createBullets(text) {
        TextBox shape = new TextBox();
        shape.text = text
        shape.anchor = new java.awt.Rectangle(50, 150, 500, 300);  //position of the text box in the slide

        RichTextRun rt = shape.textRun.richTextRuns[0];
        rt.fontSize = 42
        rt.bullet = true
        rt.bulletOffset = 50
        rt.textOffset = 100

        return shape
    }

    /**
     * Creates a shape with bullets
     */
    private def createTextbox(text) {
        TextBox shape = new TextBox();
        shape.text = text
        shape.anchor = new java.awt.Rectangle(50, 150, 500, 300);  //position of the text box in the slide

        RichTextRun rt = shape.textRun.richTextRuns[0];
        rt.fontSize = 42
        rt.textOffset = 100

        return shape
    }

    /**
     * Persist the actual PowerPoint file.
     */
    private def writeFile(SlideShow slideShow, String fileName) {
        FileOutputStream out = new FileOutputStream(fileName);
        slideShow.write(out);
        out.close();
    }

}