var office  = require('officegen');
var fs      = require('fs');
var http    = require('http');
var xml2js  = require('xml2js');
var parse   = new xml2js.Parser();

var pptgen = function() {
  var pptx = office("pptx");
  return {
    header: function(slide) {
        var el = this;
        var tmp = "header.pml";
        var data = el.readTemp(tmp);
        // el.render(slide, [{"text": data}]);
    },
    footer: function() {
       var el = this;
       var tmp = "footer.pml";
       var data = el.readTemp(tmp);
       // call render
    },
    slide: function() {
        var slide = pptx.makeNewSlide();
        return slide;
    },
    shape: function(slide, shape, opts) {
        slide.addShape(shape, opts);
        return slide;
    },
    image: function(slide, img, opts) {
        slide.addImage(img, opts);
        return slide;
    },
    text: function(slide, text, opts) {
        slide.addText(text, opts);
        return slide;
    },
    readTemp: function(file) {
      var templates = __dirname + "/templates/";
      var f = fs.readFileSync(templates + file);
      var result;
      parse.parseString(f,
        function(err, r) {
          result = r;
      });
      return result;
    },
    slides: function(slides) {
      var el = this;
      var s = slides[0].slide;
      for(var i = 0; i < s.length; i++) {
        var slide = s[i];
        var tmpSlide = el.slide();
        //check header
        // el.header(tmpSlide);
        //render content
        el.render(tmpSlide, slide.content);
        //check footer
      }
    },
    render: function(slide, data) {
      var el = this;
      //will check for content
      // console.log(data[0]);
      if(data[0].text) {
        var texts = data[0].text;
        for(var t = 0; t < texts.length; t++) {
          var text = texts[t];
          if(text['$']) {
            var txt = text["_"];
            var opts = text['$'];
            el.text(slide, txt, opts);
          }
          else {
            el.text(slide, text)
          }
        }
      }
      if(data[0].image) {
        var images = data[0].image;
        for(var t = 0; t < images.length; t++) {
          var image = images[t];
          if(image['$']) {
            var opts = image['$'];
            var src = opts.src;
            el.image(slide, src, opts);
          }
        }
      }
    },
    setOpts: function() {
      var el = this;
    },
    presentation: function() {
      var el = this;
      var tmp = "template.pml";
      var opts;
      var data = el.readTemp(tmp);
      var p = data.presentation;
      var opts = p.options[0];
      if(opts.title) {
        var title = opts.title[0];
        pptx.setDocTitle(title);
      }
      if(p.slides) {
        el.slides(p.slides);
      }
    },
    generate: function(res) {
      var el = this;
      el.presentation();
      return (res) ? pptx.generate(res) : true;
    }
  };
};



http.createServer(function(req, res) {
  var ct = "application/vnd.openxmlformats";
  ct += "-officedocument.presentationml.presentation";
  res.writeHead(200, {
    'Content-Type': ct,
    'Content-disposition': "attachment; filename=surprise.pptx"
  });
  var powerpoint = new pptgen();
  powerpoint.generate(res);
}).listen(8000);

console.log("Local server listening on port: 8000");
