/**
 *creates a scrolling background
 */
function BackGround(canvas, imgId, x, y, dx, dy){
	var bg = this;
	this.canvas = canvas;
	this.canvas2D = canvas.getContext("2d");
	this.image = document.getElementById(imgId);
	this.x = x;
	this.y = y;
	this.dx = dx;
	this.dy = dy;
	


	this.draw = function(){
		this.canvas2D.drawImage(this.image, this.x, this.y, this.canvas.width, this.canvas.height);
	};
	
	this.clear = function(){
		this.canvas2D.clearRect(0, 0, this.canvas.width, this.canvas.height);
	};
	
	this.update = function(){
		this.x += dx;
		this.y += dy;
		this.clear();
		this.draw();
	}

}
