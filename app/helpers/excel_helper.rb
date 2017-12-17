module ExcelHelper
  def apply_style(ws, *styles)
    style = {}
    styles.each do |s|
      style.merge! s
    end

    ws.styles.add_style style
  end

  def border_style
    { border: { style: :thin, color: 'FF000000' } }
  end

  def header_style
    {
      alignment: {
        horizontal: :center,
        vertical:   :center,
        wrap_text:  true
      }
    }
  end
end
