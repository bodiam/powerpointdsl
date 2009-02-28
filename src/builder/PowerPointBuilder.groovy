package builder

import exporter.SlideShowExporter

/**
 * @author Erik Pragt
 */

class PowerPointBuilder extends BuilderSupport {
    private static final String VALUE = "_value"

    SlideShow slideShow

    void setParent(parent, child) {
       println "set parent : " + parent + "," + child
    }


    def createNode(name) {
        return null
    }

    def createNode(name, Map attributes, Object value) {
        return null
    }

    def createNode(name, Object value) {
        // Only text is currently supported
        println "create node [${name}, value]"

        createNode(name, [VALUE:value])
    }

    def createNode(name, Map attributes) {
        println "create node [${name}, attributes]"

        "add${name[0].toUpperCase() + name[1..-1]}"(attributes)
    }

    def addSlideshow(attributes) {
        slideShow = new SlideShow()
        slideShow.fileName = attributes.filename ?: 'PowerPoint.ppt'

        return 'buildnode'
    }

    def addSlide(attributes) {
        slideShow.slides << new Slide(title: attributes.title)

        'slide'
    }

    def addImageslide(attributes) {
        slideShow.slides << new ImageSlide(src: attributes.src)

        'imageslide'
    }

    def addBullet(attributes) {
        slideShow.currentSlide.bullets << new Bullet(text:attributes.text)

        'bullet'
    }

    def addTextbox(attributes) {
        slideShow.currentSlide.text = attributes.VALUE

        'textbox'
    }



    void nodeCompleted(parent, node) {
        println "node completed ${parent} ${node}"

        if (node == 'buildnode') {
            new SlideShowExporter().export(slideShow)
        }
    }
}


class SlideShow {
    String fileName
    List slides = []

    Slide getCurrentSlide() {
        slides[slides.size()-1]
    }
}

class ImageSlide {
    String src
}

class Slide {
    String title
    String text
    List bullets = []
}

class Bullet {
    String text
}


