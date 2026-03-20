use std::{env, fs::File, path::PathBuf};

use image::imageops::FilterType;

fn set_icon_pixel(
    rgba: &mut [u8],
    width: usize,
    height: usize,
    x: usize,
    y: usize,
    color: [u8; 4],
) {
    if x >= width || y >= height {
        return;
    }
    let idx = (y * width + x) * 4;
    rgba[idx..idx + 4].copy_from_slice(&color);
}

fn fill_icon_rect(
    rgba: &mut [u8],
    width: usize,
    height: usize,
    x0: usize,
    y0: usize,
    x1: usize,
    y1: usize,
    color: [u8; 4],
) {
    for y in y0.min(height)..y1.min(height) {
        for x in x0.min(width)..x1.min(width) {
            set_icon_pixel(rgba, width, height, x, y, color);
        }
    }
}

fn fill_icon_circle(
    rgba: &mut [u8],
    width: usize,
    height: usize,
    cx: i32,
    cy: i32,
    radius: i32,
    color: [u8; 4],
) {
    let r2 = radius * radius;
    for y in (cy - radius).max(0)..=(cy + radius).min(height as i32 - 1) {
        for x in (cx - radius).max(0)..=(cx + radius).min(width as i32 - 1) {
            let dx = x - cx;
            let dy = y - cy;
            if dx * dx + dy * dy <= r2 {
                set_icon_pixel(rgba, width, height, x as usize, y as usize, color);
            }
        }
    }
}

fn fill_icon_ring(
    rgba: &mut [u8],
    width: usize,
    height: usize,
    cx: i32,
    cy: i32,
    outer_radius: i32,
    inner_radius: i32,
    color: [u8; 4],
) {
    let outer2 = outer_radius * outer_radius;
    let inner2 = inner_radius * inner_radius;
    for y in (cy - outer_radius).max(0)..=(cy + outer_radius).min(height as i32 - 1) {
        for x in (cx - outer_radius).max(0)..=(cx + outer_radius).min(width as i32 - 1) {
            let dx = x - cx;
            let dy = y - cy;
            let dist2 = dx * dx + dy * dy;
            if dist2 <= outer2 && dist2 >= inner2 {
                set_icon_pixel(rgba, width, height, x as usize, y as usize, color);
            }
        }
    }
}

fn draw_icon_handle(
    rgba: &mut [u8],
    width: usize,
    height: usize,
    x: usize,
    y: usize,
    length: usize,
    thickness: usize,
    color: [u8; 4],
) {
    for offset in 0..length {
        fill_icon_rect(
            rgba,
            width,
            height,
            x + offset,
            y + offset,
            x + offset + thickness,
            y + offset + thickness,
            color,
        );
    }
}

fn app_icon_rgba(width: usize, height: usize) -> Vec<u8> {
    let mut rgba = vec![0_u8; width * height * 4];

    let folder = [224, 174, 84, 255];
    let folder_top = [244, 200, 111, 255];
    let folder_shade = [196, 145, 59, 255];
    let glass = [111, 184, 236, 255];
    let glass_center = [226, 244, 255, 255];
    let outline = [86, 112, 133, 255];
    let handle = [111, 184, 236, 255];

    let px = |v: f32, axis: usize| ((v * axis as f32).round() as usize).min(axis);
    fill_icon_rect(
        &mut rgba,
        width,
        height,
        px(0.18, width),
        px(0.22, height),
        px(0.44, width),
        px(0.33, height),
        folder_top,
    );
    fill_icon_rect(
        &mut rgba,
        width,
        height,
        px(0.15, width),
        px(0.29, height),
        px(0.69, width),
        px(0.57, height),
        folder,
    );
    fill_icon_rect(
        &mut rgba,
        width,
        height,
        px(0.15, width),
        px(0.51, height),
        px(0.69, width),
        px(0.57, height),
        folder_shade,
    );
    fill_icon_rect(
        &mut rgba,
        width,
        height,
        px(0.15, width),
        px(0.28, height),
        px(0.69, width),
        px(0.38, height),
        outline,
    );

    let cx = px(0.60, width) as i32;
    let cy = px(0.52, height) as i32;
    let outer = px(0.14, width.min(height)) as i32;
    let inner = px(0.08, width.min(height)) as i32;
    fill_icon_ring(&mut rgba, width, height, cx, cy, outer, inner, glass);
    fill_icon_circle(&mut rgba, width, height, cx, cy, inner - 1, glass_center);
    draw_icon_handle(
        &mut rgba,
        width,
        height,
        px(0.66, width),
        px(0.58, height),
        px(0.13, width),
        px(0.05, width.max(height)).max(2),
        handle,
    );

    rgba
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    if env::var("CARGO_CFG_WINDOWS").is_ok() {
        let manifest_dir = PathBuf::from(env::var("CARGO_MANIFEST_DIR")?);
        let ico_icon_path = manifest_dir.join("assets").join("icons").join("app.ico");
        let png_icon_path = manifest_dir.join("assets").join("icons").join("175513.png");

        let icon_path = if ico_icon_path.is_file() {
            ico_icon_path
        } else {
            let out_dir = PathBuf::from(env::var("OUT_DIR")?);
            let generated_icon_path = out_dir.join("file-indexer.ico");
            let mut icon_dir = ico::IconDir::new(ico::ResourceType::Icon);
            if png_icon_path.is_file() {
                let source = image::open(&png_icon_path)?.to_rgba8();
                for size in [16_u32, 24, 32, 48, 64, 128, 256] {
                    let resized =
                        image::imageops::resize(&source, size, size, FilterType::Lanczos3);
                    let image = ico::IconImage::from_rgba_data(size, size, resized.into_raw());
                    icon_dir.add_entry(ico::IconDirEntry::encode(&image)?);
                }
            } else {
                for size in [16_u32, 24, 32, 48, 64, 128, 256] {
                    let image = ico::IconImage::from_rgba_data(
                        size,
                        size,
                        app_icon_rgba(size as usize, size as usize),
                    );
                    icon_dir.add_entry(ico::IconDirEntry::encode(&image)?);
                }
            }
            let mut file = File::create(&generated_icon_path)?;
            icon_dir.write(&mut file)?;
            generated_icon_path
        };

        winresource::WindowsResource::new()
            .set_icon(icon_path.to_string_lossy().as_ref())
            .compile()?;
    }

    Ok(())
}
