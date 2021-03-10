function px_to_pt(size: string) {
  let len = size.length;
  if (size.substr(len-2) != "px") {
      return 0;
  } else {
      const v = parseFloat(size.substring(0, len-2))
      return isNaN(v) ? 0 : v * 0.75
  }
}


function color_to_argb(color: string) {
  const len = color.length
  let argb_color = ''
  if (len === 7 && color[0] == "#") {
      argb_color += "FF"
      argb_color += color.substr(1)
      return argb_color
  } else if (len > 11 && color.substr(0, 5) == "rgba(" && color.substr(len - 1) == ")") {
      const colors_part = color.substring(5, len-1)
      const colors = colors_part.split(",").map((s) => s.trim())
      if (colors.length < 4) {
          return ''
      }
      const r = str_to_hex(colors[0]);
      const g = str_to_hex(colors[1]);
      const b = str_to_hex(colors[2]);
      const a = str_alpha_to_hex(colors[3]);
      if (!r || !g || !b || !a) {
          return '';
      }
      argb_color += `${a}${r}${g}${b}`;
      return argb_color;
  } else if (len > 10 && color.substr(0, 4) == "rgb(" && color.substr(len-1) == ")") {
      const colors_part = color.substring(4, len-1);
      const colors = colors_part.split(",").map((s) => s.trim());
      if (colors.length < 3) {
          return '';
      }
      let r = str_to_hex(colors[0]);
      let g = str_to_hex(colors[1]);
      let b = str_to_hex(colors[2]);
      if (!r || !g || !b) {
          return '';
      }
      argb_color += `FF${r}${g}${b}`
      return argb_color
  }
  return ''
}

function str_to_hex(s: string) {
  const v = parseInt(s)
  if (!isNaN(v)) {
    const res = v.toString(16).toUpperCase()
    if (res.length === 1) {
      return `0${res}`
    } else if (res.length === 2) {
      return res
    }
  }
  return ''
}

function str_alpha_to_hex(s: string) {
  const v = parseInt(s)
  const res = (v * 255).toString(16).toUpperCase()
  if (res.length === 1) {
    return `0${res}`
  } else if (res.length === 2) {
    return res
  }
  return ''
}

function setDefaultProp(obj: Record<string, any>, key: string, defaultValue: any) {
  if (!(key in obj)) {
    obj[key] = defaultValue
  }
  return obj
}

export function handleStyle(style: Record<string, string>) {
  const ret: Record<string, any> = {}
  for (const key in style) {
    switch (key) {
      case 'color':
        setDefaultProp(ret, 'font', {})
        ret.font.color = {
          argb: color_to_argb(style[key])
        }
        break;
      case 'fontSize':
        setDefaultProp(ret, 'font', {})
        ret.font.size = px_to_pt(style[key])
        break;
      case 'fontWeight':
        setDefaultProp(ret, 'font', {})
        ret.font.bold = true
        break;
      case 'verticalAlign':
        setDefaultProp(ret, 'alignment', {})
        ret.alignment.vertical = style[key] === 'center' ? 'middle' : style[key]
        break
      case 'background':
        setDefaultProp(ret, 'fill', {})
        ret.fill.type = 'pattern'
        ret.fill.pattern = 'solid'
        ret.fill.bgColor = { argb: color_to_argb(style[key]) }
        ret.fill.fgColor = ret.fill.bgColor
        break;
      case 'textAlign':
        setDefaultProp(ret, 'alignment', {})
        ret.alignment.horizontal = style[key]
    }
  }
  return ret
}
