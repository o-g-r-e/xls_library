import java.awt.Color;

public class ExcelStyle {
	private Color backgroundColor;
	private Color borderColor;
	private ExcelFont font;
	
	protected ExcelStyle() {}
	
	public Color getBackgroundColor() {
		return backgroundColor;
	}
	public void setBackgroundColor(Color backgroundColor) {
		this.backgroundColor = backgroundColor;
	}
	public Color getBorderColor() {
		return borderColor;
	}
	public void setBorderColor(Color borderColor) {
		this.borderColor = borderColor;
	}
	public ExcelFont getFont() {
		return font;
	}
	public void setFont(ExcelFont font) {
		this.font = font;
	}
}
