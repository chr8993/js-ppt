var office      = require('officegen');
var fs          = require('fs');
var xml2js      = require('xml2js');
var parse       = new xml2js.Parser();
var Handlebars  = require('handlebars');

var pptgen = function(opts) {
  var args = arguments;
  var bindData = (opts.data) ? opts.data: {};
  var contentTemplate = (opts.template) ? opts.template: "";
  var headerTemplate = (opts.header) ? opts.header: "";
  var footerTemplate = (opts.footer) ? opts.footer: "";
  var set = {
    type: "pptx"
  };
  var pptx = office(set);
  return {
    header: function(slide) {
        if(headerTemplate) {
          var el = this;
          var tmp = headerTemplate;
          var data = el.readTemp(tmp);
          var cnt = [data.content];
          el.render(slide, cnt);
        }
    },
    footer: function(slide) {
        if(footerTemplate) {
          var el = this;
          var tmp = "footer.pml";
          var data = el.readTemp(tmp);
          var cnt = [data.content];
          el.render(slide, cnt);
        }
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
    readTemp: function(filepath) {
      var f = String(fs.readFileSync(filepath));
      var result;
      if(bindData) {
        var a = bindData;
        var tmp = Handlebars.compile(f);
        var compiled = tmp(a);
        f = compiled;
      }
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
        el.header(tmpSlide);
        el.render(tmpSlide, slide.content);
        el.footer(tmpSlide);
      }
    },
    render: function(slide, data) {
      var el = this;
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
    },
    setOpts: function() {
      var el = this;
    },
    presentation: function() {
      var el = this;
      var tmp = contentTemplate;
      if(!tmp) return;
      var pres = el.readTemp(tmp);
      var p = pres.presentation;
      var options = p.options[0];
      if(options.title) {
        var title = options.title[0];
        pptx.setDocTitle(title);
      }
      if(p.slides) {
        el.slides(p.slides);
      }
    },
    generate: function(res) {
      var el = this;
      el.presentation();
      pptx.generate(res);
    }
  };
};

module.exports = pptgen;
