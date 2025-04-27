VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SVGmodel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Const svgNS   As String = "http:''www.w3.org/2000/svg"

Dim sizeCache  As New Dictionary '= {}

Public svg As Dictionary     ' SVGElement
Public currentGroup As iArray      'SVGElement[]
Public dummySvg
Public Sub init() 'wrapper)
        Set Me.svg = createSvg()
        Set Me.currentGroup = New iArray
        'line wrapper.appendChild Me.svg
End Sub

Public Sub Clear()
'        Dim wrapper
'        If Not (Me.svg Is Nothing) Then
'            wrapper = Me.svg.parentNode
'            Me.svg = createSvg()
'            Me.currentGroup = []
'            if (wrapper) {
'                '' TODO-PER: If the wrapper is not present, then the underlying div was pulled out from under this instance. It's possible that is still useful (for creating the music off page?)
'                wrapper.innerHTML = ""
'                wrapper.appendChild (Me.svg)
'            End If
'        End If
End Sub

'    setTitle(title) {
'        var titleEl = document.createElement("title")
'        var titleNode = document.createTextNode(title)
'        titleEl.appendChild(titleNode)
'        me.svg.insertBefore(titleEl, me.svg.firstChild)
'    }
'
'    setResponsiveWidth(w, h) {
'        '' this technique is from: http:''thenewcode.com/744/Make-SVG-Responsive, thx to https:''github.com/iantresman
'        me.svg.setAttribute("viewBox", "0 0 " + w + " " + h)
'        me.svg.setAttribute("preserveAspectRatio", "xMinYMin meet")
'        me.svg.removeAttribute("height")
'        me.svg.removeAttribute("width")
'        me.svg.style['display'] = "inline-block"
'        me.svg.style['position'] = "absolute"
'        me.svg.style['top'] = "0"
'        me.svg.style['left'] = "0"
'
'        if (me.svg.parentNode) {
'            var cls = me.svg.parentNode.getAttribute("class")
'            if (!cls)
'                me.svg.parentNode.setAttribute("class", "abcjs-container")
'            else if (cls.indexOf("abcjs-container") < 0)
'                me.svg.parentNode.setAttribute("class", cls + " abcjs-container")
'            me.svg.parentNode.style['display'] = "inline-block"
'            me.svg.parentNode.style['position'] = "relative"
'            me.svg.parentNode.style['width'] = "100%"
'            '' PER: I changed the padding from 100% to this through trial and error.
'            '' The example was using a square image, but this music might be either wider or taller.
'            var padding = h / w * 100
'            me.svg.parentNode.style['padding-bottom'] = padding + "%"
'            me.svg.parentNode.style['vertical-align'] = "middle"
'            me.svg.parentNode.style['overflow'] = "hidden"
'        }
'    }
'
'    setSize(w, h) {
'        me.svg.setAttribute('width', w)
'        me.svg.setAttribute('height', h)
'    }
'
'    setAttribute(attr, value) {
'        me.svg.setAttribute(attr, value)
'    }
'
'    setScale(scale) {
'        if (scale !== 1) {
'            me.svg.style.transform = "scale(" + scale + "," + scale + ")"
'            me.svg.style['-ms-transform'] = "scale(" + scale + "," + scale + ")"
'            me.svg.style['-webkit-transform'] = "scale(" + scale + "," + scale + ")"
'            me.svg.style['transform-origin'] = "0 0"
'            me.svg.style['-ms-transform-origin-x'] = "0"
'            me.svg.style['-ms-transform-origin-y'] = "0"
'            me.svg.style['-webkit-transform-origin-x'] = "0"
'            me.svg.style['-webkit-transform-origin-y'] = "0"
'        } else {
'            me.svg.style.transform = ""
'            me.svg.style['-ms-transform'] = ""
'            me.svg.style['-webkit-transform'] = ""
'        }
'    }
'
'    insertStyles(styles) {
'        var el = document.createElementNS(svgNS, "style")
'        el.textContent = styles
'        me.svg.insertBefore(el, me.svg.firstChild)  '' prepend is not available on older browsers.
'        ''  me.svg.prepend(el)
'    }
'
'    setParentStyles(attr) {
'        '' This is needed to get the size right when there is scaling involved.
'        for (var key in attr) {
'            if (attr.hasOwnProperty(key)) {
'                if (me.svg.parentNode)
'                    me.svg.parentNode.style[key] = attr[key]
'            }
'        }
'        '' This is the last thing that gets called, so delete the temporary SVG if one was created
'        if (me.dummySvg) {
'            var body = document.querySelector('body')
'            body.removeChild(me.dummySvg)
'            me.dummySvg = null
'        }
'
'    }
'
'
'    rect(attr) {
'        '' This uses path instead of rect so that it can be hollow and the color changes with "fill" instead of "stroke".
'        var lines = []
'        var x1 = attr.x
'        var y1 = attr.y
'        var x2 = attr.x + attr.width
'        var y2 = attr.y + attr.height
'        lines.push(constructHLine(x1, y1, x2))
'        lines.push(constructHLine(x1, y2, x2))
'        lines.push(constructVLine(x2, y1, y2))
'        lines.push(constructVLine(x1, y2, y1))
'
'        return me.path({ path: lines.join(" "), stroke: "none", "data-name": attr["data-name"] })
'    }
'
'    dottedLine(attr) {
'        var el = document.createElementNS(svgNS, 'line')
'        el.setAttribute("x1", attr.x1)
'        el.setAttribute("x2", attr.x2)
'        el.setAttribute("y1", attr.y1)
'        el.setAttribute("y2", attr.y2)
'        el.setAttribute("stroke", attr.stroke)
'        el.setAttribute("stroke-dasharray", "5,5")
'        me.svg.insertBefore(el, me.svg.firstChild)
'    }
'
'    rectBeneath(attr) {
'        var el = document.createElementNS(svgNS, 'rect')
'        el.setAttribute("x", attr.x)
'        el.setAttribute("width", attr.width)
'        el.setAttribute("y", attr.y)
'        el.setAttribute("height", attr.height)
'        if (attr.stroke)
'            el.setAttribute("stroke", attr.stroke)
'        if (attr['stroke-opacity'])
'            el.setAttribute("stroke-opacity", attr['stroke-opacity'])
'        if (attr.fill)
'            el.setAttribute("fill", attr.fill)
'        if (attr['fill-opacity'])
'            el.setAttribute("fill-opacity", attr['fill-opacity'])
'        me.svg.insertBefore(el, me.svg.firstChild)
'    }
'
'    text(text, attr, target?): SVGTextElement {
'        var el = document.createElementNS(svgNS, 'text') as SVGTextElement
'        el.setAttribute("stroke", "none")
'        for (var key in attr) {
'            if (attr.hasOwnProperty(key)) {
'                el.setAttribute(key, attr[key])
'            }
'        }
'        var lines = ("" + text).split("\n")
'        for (var i = 0  i < lines.length  i++) {
'            var line = document.createElementNS(svgNS, 'tspan')
'            line.setAttribute("x", attr.x ? attr.x : 0)
'            if (i !== 0)
'                line.setAttribute("dy", "1.2em")
'            if (lines[i].indexOf("\x03") !== -1) {
'                var parts = lines[i].split('\x03')
'                line.textContent = parts[0]
'                if (parts[1]) {
'                    var ts2 = document.createElementNS(svgNS, 'tspan')
'                    ts2.setAttribute("dy", "-0.3em")
'                    ts2.setAttribute("style", "font-size:0.7em")
'                    ts2.textContent = parts[1]
'                    line.appendChild(ts2)
'                }
'                if (parts[2]) {
'                    var dist = parts[1] ? "0.4em" : "0.1em"
'                    var ts3 = document.createElementNS(svgNS, 'tspan')
'                    ts3.setAttribute("dy", dist)
'                    ts3.setAttribute("style", "font-size:0.7em")
'                    ts3.textContent = parts[2]
'                    line.appendChild(ts3)
'                }
'            } else
'                line.textContent = lines[i]
'            el.appendChild(line)
'        }
'        if (target)
'            target.appendChild(el)
'        Else
'            me.append(el)
'        return el
'    }
'
'    guessWidth(text, attr): { width: number height:number} {
'        var svg = me.createDummySvg()
'        var el = me.text(text, attr, svg)
'        var size: { width: number  height: number }
'        try {
'            size = el.getBBox()
'            if (isNaN(size.height) || !size.height) '' TODO-PER: I don't think this can happen unless there isn't a browser at all.
'                size = { width: attr['font-size'] / 2, height: attr['font-size'] + 2 }  '' Just a wild guess.
'            Else
'                size = { width: size.width, height: size.height }
'        } catch (ex) {
'            size = { width: attr['font-size'] / 2, height: attr['font-size'] + 2 }  '' Just a wild guess.
'        }
'        svg.removeChild(el)
'        return size
'    }
'
'    createDummySvg(): SVGElement {
'        if (!me.dummySvg) {
'            me.dummySvg = createSvg()
'            var styles = [
'                "display: block !important ",
'                "height: 1px ",
'                "width: 1px ",
'                "position: absolute "
'            ]
'            me.dummySvg.setAttribute('style', styles.join(""))
'            var body = document.querySelector('body')
'            body.appendChild(me.dummySvg)
'        }
'
'        return me.dummySvg
'    }
'
'
Public Function getTextSize(text, attr, el As Dictionary)   ': SVGTextElement) {

'        if (typeof text === 'number')
'            text = '' + text
'        if (!text || text.match(/^\s+$/))
'            return { width: 0, height: 0 }
'        var key
'        if (text.length < 20) {
'            '' The short text tends to be repetitive and getBBox is really slow, so lets cache.
'            key = text + JSON.stringify(attr)
'            if (sizeCache[key])
'                return sizeCache[key]
'        }
'        var removeLater = !el
'        if (!el)
'            el = me.text(text, attr)
'        var size
'        try {
'            size = el.getBBox()
'            if (isNaN(size.height) || !size.height)
'                size = me.guessWidth(text, attr)
'            Else
'                size = { width: size.width, height: size.height }
'        } catch (ex) {
'            size = me.guessWidth(text, attr)
'        }
'        if (removeLater) {
'            if (me.currentGroup.length > 0)
'                me.currentGroup[0].removeChild(el)
'            Else
'                me.svg.removeChild(el)
'        }
'        if (key)
'            sizeCache[key] = size
'        return size

End Function
'
'    openGroup(options?) {
'        options = options ? options : {}
'        var el = document.createElementNS(svgNS, "g") as SVGGElement
'        if (options.klass)
'            el.setAttribute("class", options.klass)
'        if (options.fill)
'            el.setAttribute("fill", options.fill)
'        if (options.stroke)
'            el.setAttribute("stroke", options.stroke)
'        if (options['data-name'])
'            el.setAttribute("data-name", options['data-name'])
'
'        if (options.prepend)
'            me.prepend(el)
'        Else
'            me.append(el)
'        me.currentGroup.unshift(el)
'        return el
'    }
'
'    closeGroup() {
'        var g = me.currentGroup.shift()
'        if (g && g.children.length === 0) {
'            '' If nothing was added to the group it is because all the elements were invisible. We don't need the group, then.
'            g.parentElement.removeChild(g)
'            return null
'        }
'        return g
'    }
'
'    path(attr): SVGPathElement {
'        var el = document.createElementNS(svgNS, "path") as SVGPathElement
'        for (var key in attr) {
'            if (attr.hasOwnProperty(key)) {
'                if (key === 'path')
'                    el.setAttributeNS(null, 'd', attr.path)
'                else if (key === 'klass')
'                    el.setAttributeNS(null, "class", attr[key])
'                else if (attr[key] !== undefined)
'                    el.setAttributeNS(null, key, attr[key])
'            }
'        }
'        me.append(el)
'        return el
'    }
'
'    pathToBack(attr): SVGPathElement {
'        var el = document.createElementNS(svgNS, "path") as SVGPathElement
'        for (var key in attr) {
'            if (attr.hasOwnProperty(key)) {
'                if (key === 'path')
'                    el.setAttributeNS(null, 'd', attr.path)
'                else if (key === 'klass')
'                    el.setAttributeNS(null, "class", attr[key])
'                Else
'                    el.setAttributeNS(null, key, attr[key])
'            }
'        }
'        me.prepend(el)
'        return el
'    }
'
'    lineToBack(attr): SVGLineElement {
'        var el = document.createElementNS(svgNS, 'line') as SVGLineElement
'        var Keys = Object.Keys(attr)
'        for (var i = 0  i < keys.length  i++)
'            el.setAttribute(keys[i], attr[keys[i]])
'        me.prepend(el)
'        return el
'    }
'
'
'    append(el: SVGElement) {
'        if (me.currentGroup.length > 0)
'            me.currentGroup[0].appendChild(el)
'        Else
'            me.svg.appendChild(el)
'    }
'
'    prepend(el: SVGElement) {
'        '' The entire group is prepended, so don't prepend the individual elements.
'        if (me.currentGroup.length > 0)
'            me.currentGroup[0].appendChild(el)
'        Else
'            me.svg.insertBefore(el, me.svg.firstChild)
'    }
'
'    setAttributeOnElement(el: SVGElement, attr) {
'        for (var key in attr) {
'            if (attr.hasOwnProperty(key)) {
'                el.setAttributeNS(null, key, attr[key])
'            }
'        }
'    }
'
'    moveElementToChild(parent, child) {
'        parent.appendChild(child)
'    }
'}
'function constructHLine(x1, y1, x2) {
'    var len = x2 - x1
'    return "M " + x1 + " " + y1 +
'        " l " + len + ' ' + 0 +
'        " l " + 0 + " " + 1 + " " +
'        " l " + (-len) + " " + 0 + " " + " z "
'}
'
'function constructVLine(x1, y1, y2) {
'    var len = y2 - y1
'    return "M " + x1 + " " + y1 +
'        " l " + 0 + ' ' + len +
'        " l " + 1 + " " + 0 + " " +
'        " l " + 0 + " " + (-len) + " " + " z "
'}
Public Function createSvg()
'
'    var svg = document.createElementNS(svgNS, "svg")
'    svg.setAttributeNS("http:''www.w3.org/2000/xmlns/", "xmlns:xlink", "http:''www.w3.org/1999/xlink")
'    svg.setAttribute('role', 'img')     '' for accessibility
'    svg.setAttribute('fill', 'currentColor')     '' for automatically picking up dark mode and high contrast
'    svg.setAttribute('stroke', 'currentColor')     '' for automatically picking up dark mode and high contrast
'    return svg as SVGElement
    Set createSvg = Nothing
End Function


