import builder.PowerPointBuilder

/**
 * @author Erik Pragt
 */

def builder = new PowerPointBuilder()

builder.slideshow(filename:'Test.ppt') {
    slide(title: 'Introduction') {
        bullet(text: 'Bullet 1')
        bullet(text: 'Bullet 2')
    }
    slide(title: 'Slide 2') {
        bullet(text: 'Bullet 3')
        bullet(text: 'Bullet 4')
    }
    slide(title: 'Example') {
        textbox("""This is a lot of text!!
With a lot of extra lines
Which make no sense
Which make no sense
Which make no sense
At all""")
    }
    imageslide(src:'background.png')
    /*
    image(src:'lala.png') {
        
    }*/
}

/*
builder.build {
  page(title:'Introduction') {
     bullet(text:'Bullet 1')
     bullet(text:'Bullet 2')
     bullet(text:'Bullet 3') {
         bullet(text:'Bullet 3.1')
         bullet(text:'Bullet 3.2')
                 bullet(text:'Bullet 3.3')
     }
     notes(text:'Tell something about the introduction')
  }
  page(title:'Questions?') {
      image('images/questions.gif')
  }
}
*/
