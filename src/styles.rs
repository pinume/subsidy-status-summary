use rust_xlsxwriter::{Color, Format, FormatAlign, FormatBorder};

/// Shared styling formats for Excel generation
pub struct StylePool {
    pub title: Format,
    pub subtitle: Format,
    pub section_header: Format,
    pub col_header: Format,
    pub col_header_center: Format,
    pub tip_row: Format,

    pub text_left: Format,
    pub text_center: Format,
    pub text_right: Format,

    pub int_count: Format,
    pub money_wan: Format,
    pub money_yuan: Format,
    pub percent: Format,

    pub date: Format,
    pub datetime: Format,
}

impl Default for StylePool {
    fn default() -> Self {
        Self::new()
    }
}

impl StylePool {
    pub fn new() -> Self {
        let border_color = Color::RGB(0xD9D9D9);
        let border_style = FormatBorder::Thin;
        let font_family = "微软雅黑";

        let title = Format::new()
            .set_font_name(font_family)
            .set_font_size(18)
            .set_bold()
            .set_font_color(Color::White)
            .set_background_color(Color::RGB(0x1F4E78))
            .set_align(FormatAlign::Center)
            .set_align(FormatAlign::VerticalCenter);

        let subtitle = Format::new()
            .set_font_name(font_family)
            .set_font_size(11)
            .set_align(FormatAlign::Left)
            .set_align(FormatAlign::VerticalCenter);

        let section_header = Format::new()
            .set_font_name(font_family)
            .set_font_size(12)
            .set_bold()
            .set_font_color(Color::White)
            .set_background_color(Color::RGB(0x1F4E78))
            .set_align(FormatAlign::Center)
            .set_align(FormatAlign::VerticalCenter)
            .set_border(border_style)
            .set_border_color(border_color);

        let col_header = Format::new()
            .set_font_name(font_family)
            .set_font_size(11)
            .set_bold()
            .set_font_color(Color::White)
            .set_background_color(Color::RGB(0x5B9BD5))
            .set_align(FormatAlign::Left)
            .set_align(FormatAlign::VerticalCenter)
            .set_border(border_style)
            .set_border_color(border_color);

        let col_header_center = Format::new()
            .set_font_name(font_family)
            .set_font_size(11)
            .set_bold()
            .set_font_color(Color::White)
            .set_background_color(Color::RGB(0x5B9BD5))
            .set_align(FormatAlign::Center)
            .set_align(FormatAlign::VerticalCenter)
            .set_border(border_style)
            .set_border_color(border_color);

        let tip_row = Format::new()
            .set_font_name(font_family)
            .set_font_size(11)
            .set_bold()
            .set_font_color(Color::Black)
            .set_background_color(Color::RGB(0xD9EAF7))
            .set_align(FormatAlign::Left)
            .set_align(FormatAlign::VerticalCenter)
            .set_border(border_style)
            .set_border_color(border_color);

        let text_left = Format::new()
            .set_font_name(font_family)
            .set_font_size(11)
            .set_align(FormatAlign::Left)
            .set_align(FormatAlign::VerticalCenter)
            .set_border(border_style)
            .set_border_color(border_color);

        let text_center = Format::new()
            .set_font_name(font_family)
            .set_font_size(11)
            .set_align(FormatAlign::Center)
            .set_align(FormatAlign::VerticalCenter)
            .set_border(border_style)
            .set_border_color(border_color);

        let text_right = Format::new()
            .set_font_name(font_family)
            .set_font_size(11)
            .set_align(FormatAlign::Right)
            .set_align(FormatAlign::VerticalCenter)
            .set_border(border_style)
            .set_border_color(border_color);

        let int_count = Format::new()
            .set_font_name(font_family)
            .set_font_size(11)
            .set_num_format("#,##0")
            .set_align(FormatAlign::Right)
            .set_align(FormatAlign::VerticalCenter)
            .set_border(border_style)
            .set_border_color(border_color);

        let money_wan = Format::new()
            .set_font_name(font_family)
            .set_font_size(11)
            .set_num_format("#,##0.00;[Red]-#,##0.00")
            .set_align(FormatAlign::Right)
            .set_align(FormatAlign::VerticalCenter)
            .set_border(border_style)
            .set_border_color(border_color);

        let money_yuan = Format::new()
            .set_font_name(font_family)
            .set_font_size(11)
            .set_num_format("#,##0.00;[Red]-#,##0.00")
            .set_align(FormatAlign::Right)
            .set_align(FormatAlign::VerticalCenter)
            .set_border(border_style)
            .set_border_color(border_color);

        let percent = Format::new()
            .set_font_name(font_family)
            .set_font_size(11)
            .set_num_format("0.00%")
            .set_align(FormatAlign::Right)
            .set_align(FormatAlign::VerticalCenter)
            .set_border(border_style)
            .set_border_color(border_color);

        let date = Format::new()
            .set_font_name(font_family)
            .set_font_size(11)
            .set_num_format("yyyy-mm-dd")
            .set_align(FormatAlign::Center)
            .set_align(FormatAlign::VerticalCenter)
            .set_border(border_style)
            .set_border_color(border_color);

        let datetime = Format::new()
            .set_font_name(font_family)
            .set_font_size(11)
            .set_num_format("yyyy-mm-dd hh:mm:ss")
            .set_align(FormatAlign::Center)
            .set_align(FormatAlign::VerticalCenter)
            .set_border(border_style)
            .set_border_color(border_color);

        Self {
            title,
            subtitle,
            section_header,
            col_header,
            col_header_center,
            tip_row,
            text_left,
            text_center,
            text_right,
            int_count,
            money_wan,
            money_yuan,
            percent,
            date,
            datetime,
        }
    }
}
