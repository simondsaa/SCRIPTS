
		//custom math library
		var BMath = {
			max : function(x, y){
				return (x > y)? x : y;
			},
			
			min : function(x, y){
				return (y > x)? x : y;				
			}
		};

		//sounds
		var bgMusic = new Audio("../TurtleWarrior/music.mp3");
		var screamSound = new Audio("../TurtleWarrior/scream-shutup.wav");
		var shot = new Audio("../TurtleWarrior/shot.mp4");
		var explosion = new Audio("../TurtleWarrior/explosion.mp3");
		var gun_refill = new Audio("../TurtleWarrior/gun_refill.mp3");
		var fuel_refill = new Audio("../TurtleWarrior/laser_goal.mp3");
		var tickSound = new Audio("../TurtleWarrior/Tick.mp3");
		var health_pick = new Audio("../TurtleWarrior/healthPick.mp3");
		health_pick.volume = "1";

window.onload = function(){

		$(".hidden").hide();
	

		var started = false;

		//Used to make sure one key is not registered twice once pressed
		var lock = 0;

		var canvas = document.getElementById("gameCanvas");
		var canvas2D = canvas.getContext("2d");


		
		var CANVAS_WIDTH = 	800;
		var CANVAS_HEIGHT = 600;
		
		var FPS = 9;
		

		//All cursors in the whole manu and in all pages		
		var cursors = [];
		//cursors in page 1
		cursors.push(new Cursor(600, [205, 300, 375], document.getElementById("cursor2")));
		cursors.push(new Cursor(200, [205, 300, 375], document.getElementById("cursor1")));

		//cursors in page 2
		cursors.push(new Cursor(250, [470], document.getElementById("cursor2")));
		cursors.push(new Cursor(50, [470], document.getElementById("cursor1")));

		//cursors in page 3
		cursors.push(new Cursor(460, [490], document.getElementById("cursor2")));
		cursors.push(new Cursor(250, [490], document.getElementById("cursor1")));

		

		var pages = [];

		//page 1
		pages.push(new Page(cursors[0], cursors[1], document.getElementById("menu")));

		//page 2
		pages.push(new Page(cursors[2], cursors[3], document.getElementById("credits")));

		//page 3
		pages.push(new Page(cursors[4], cursors[5], document.getElementById("instructions")));


		var currentPage = 0;



		canvas.addEventListener("mousedown", function(event){
			console.log("[" + event.pageX + ", " + event.pageY + "]");
		}, false);

		

		
		var intervalID = setInterval(function(){
			if(!started){
				update();
				draw();
			}
		}, 1000/FPS);


		function draw(){
			canvas2D.clearRect(0, 0, 800, 600);
			
			//draw current page
			canvas2D.drawImage(pages[currentPage].image, 0, 0, 800, 550);

			//draw cursors on current page
			for(var i=0; i<2; i++)
				pages[currentPage].cursors[i].draw();

			
		}
		
		function update(){

			for(var i=0; i<pages[currentPage].cursors.length; i++){
				cursors[i].update();
			}						
		}

		function Page(cursor1, cursor2, img){
			this.cursors = [cursor1, cursor2];
			this.image = img;
			
		}

		//CLASSES
		function Cursor(x, yPosArr,img){
			
			this.x = x;
			this.image = img;
			this.yPosisionsArr = yPosArr;
			this.currentPosision = 0;


			this.draw = function(){
				canvas2D.drawImage(this.image, this.x, this.yPosisionsArr[this.currentPosision], 40, 40);
			};
			
			this.update = function(){
				//console.log(currentPage);
				
				//If we are on the first page
				if(currentPage == 0){
					if(keydown.down){
						tickSound.pause();
						tickSound.currentTime = 0;
						tickSound.play();
						this.currentPosision = (this.currentPosision + 1)  % 3;
					}else if(keydown.up){
						tickSound.pause();
						tickSound.currentTime = 0;
						tickSound.play();
						this.currentPosision = (this.currentPosision - 1)  % 3;
						if(this.currentPosision == -1){
							this.currentPosision = 2;
						}				
					}else if((keydown.space  || keydown.return)){
						if(this.currentPosision == 0){
							clearInterval(intervalID);
							restart();
						}else if(this.currentPosision == 1){
							currentPage = 1;
							lock = 1 - lock;
	
						}else if(this.currentPosision == 2){
							currentPage = 2;
							lock = 1 - lock;
						}
						
					}
				//If we are on the credits (second) page
				}else if(currentPage == 1){
					if((keydown.space  || keydown.return) && lock != 1){
						currentPage = 0;
					}else{
						lock = 1 - lock;
					}
					
				}else if(currentPage == 2){
					if((keydown.space  || keydown.return) && lock != 1){
						currentPage = 0;
					}else{
						lock = 1 - lock;
					}
				}
				
			};
			
		}

	}


	function restart(){

		
		console.log(BMath.min(10, 11));

		var canvas = document.getElementById("gameCanvas");
		var canvas2D = canvas.getContext("2d");
		
		var CANVAS_WIDTH = 	800;
		var CANVAS_HEIGHT = 600;
		
		var FPS = 24;

	
		var bulletsLeft = 100;
		var fuelLeft = 300;
		var health = 100;

		var username = "";
	
		var score = 0;
	
		var paused = false;
	

		var GRAVITY = 5;
		var img = document.getElementById("charImage");
		var img2 = document.getElementById("charImage2");
		var bgImg = document.getElementById("bgImage");
		var jumpSound = new Audio("../TurtleWarrior/jump.mp3"); // buffers automatically when created
		var bullets = []; 
		var bubbles = [];
		var enemies= [];
		var bonuses = [];

	
		var cHandler = new CollisionHandler();
	

		canvas.addEventListener("mousedown", function(event){
			console.log("[" + event.pageX + ", " + event.pageY +  "]");
		}, false);		
	
		bgMusic.volume = 0.3;
		bgMusic.play();
	      
		setInterval(update, 1000/FPS);
		setInterval(draw, 1000/FPS);
		setInterval(function(){
			if(!paused)
				var rand = Math.random();
				enemies.push(new Enemy(800, rand*100 + 400*(1 - rand), -Math.random()*5 - 2, Math.round(Math.random()*4)));
		}, 250);
	
		setInterval(function(){
			if(!paused)
				bonuses.push(new Bonus(Math.random()*800, 0, Math.random()*10, Math.round(Math.random()*3)));
				//console.log(bonuses[bonuses.length -1]);
		}, 5000);


	
		function pause(){
			for(var i=0; i<enemies.length; i++){
				enemies[i].dx = 0;
			}
	
			for(var i=0; i<bullets.length; i++){
				bullets[i].dx = 0;
			}
			
			player.dx = 0;
			player.dy = 0;
				
				
		}
	
		function update(){
			player.move(); 
	
			for(var i=0; i<bullets.length; i++){
				if(bullets[i].x > CANVAS_WIDTH || bullets[i].x < 0)
					bullets.splice(i, 1);
				else
					bullets[i].move();

				for(var j=0; j<enemies.length; j++){
					if(cHandler.isCollision(enemies[j], bullets[i])){
						bullets.splice(i, 1);
						enemies[j].health--;
						explosion.pause();
						explosion.currentTime = 0;
						explosion.play();
						if(enemies[j].health < 1){
							enemies.splice(j, 1);
							score += enemies[j].type;
						}
					}
				}
			}
			
			for(var i=0; i<bonuses.length; i++){
				bonuses[i].move();
				if(cHandler.isCollision(player, bonuses[i])){
					if(bonuses[i].type == 1){
						fuel_refill.pause();
						fuel_refill.currentTime = 0;
						fuel_refill.play();
						fuelLeft = BMath.min(fuelLeft + 100, 300);					
					}else if(bonuses[i].type == 2){
						gun_refill.pause();
						gun_refill.currentTime = 0;
						gun_refill.play();
						bulletsLeft += 300;
					}else if(bonuses[i].type == 3){
						health_pick.pause();
						health_pick.currentTime = 0;
						health_pick.play();
						health = BMath.min(health + 20, 100);
					}
					bonuses.splice(i, 1);

				}
				
			}
			
			for(var i=0; i<enemies.length; i++){
				enemies[i].move();			
				if(enemies[i].x < 1){
					enemies.splice(i, 1);
					if(health > 0)
						health -= 1;
					else if(!paused){
						bgMusic.pause();
						paused = true;
						document.getElementById("highScore").innerHTML = "High Score: " + score;
					}
				}
	
			}
	
			for(var i=0; i<bubbles.length; i++){
				if(bubbles[i].y - bubbles[i].yInit > 1)
					bubbles.splice(i, 1);
				else
					bubbles[i].move();
			}
	
	
	
	
		}
		
		function draw(){
			if(paused){
				canvas2D.font = '24pt Calibri';			
				canvas2D.drawImage(document.getElementById("gameOver"), 0, 0, 800, 600);
				canvas2D.fillText("You scored " + score + " " + username, 311, 500);
				pause();
				return;
			}
			
			canvas2D.fillStyle = "black";
			canvas2D.clearRect(0, 0, CANVAS_WIDTH, CANVAS_HEIGHT);
			canvas2D.drawImage(bgImg, 0, 0, CANVAS_WIDTH, CANVAS_HEIGHT - 50);
			canvas2D.font = 'italic 12pt Calibri';
			canvas2D.fillText("Developed by Behnam Azizi", 600, 530);
	
			canvas2D.fillText("Score: " + score, 100, 20);
	
	
			canvas2D.fillText("Bullets: " + bulletsLeft, 300, 20);
	
			canvas2D.fillText("Jet Pack: ", 400, 20);
			canvas2D.fillRect(460, 5, fuelLeft/2, 20);
	
			canvas2D.fillText("Health: ", 650, 20);
			canvas2D.fillStyle = "red";
			canvas2D.fillRect(700, 5, health/2, 20);
	
			
			player.draw();
	
			for(var i=0; i<bullets.length; i++)
				bullets[i].draw();
	
			for(var i=0; i<bubbles.length; i++)
				bubbles[i].draw();
	
			for(var i=0; i<enemies.length; i++)
				enemies[i].draw();

			for(var i=0; i<bonuses.length; i++)
				bonuses[i].draw();
	
	
		}
	
		/*---------------------------   Game objects ---------------------------
		 * 
		 * 
		 */
		
		var ground = {
			x: 0,
			y:500,
			width:800,
			height: 600
		};
		
		function Bullet(x, y, face){
			this.image = document.getElementById("bulletImage");
			this.dyInit = 5;
			this.dx = 20;
			if(face == "left")
				this.dx = -this.dx;
			this.x = x;
			this.y = y;
			this.width = 20;
			this.height = 10
			this.move = function(){
				this.x += this.dx;			
			};
			
			this.draw = function(){
				if(face == "right")
					this.image = document.getElementById("bulletImage2");
				else
					this.image = document.getElementById("bulletImage");
					
				canvas2D.drawImage(this.image, this.x, this.y, this.width, this.height);
			}
			
		}
		
		var player = {
			//color: "#000",
			jumpHeight: 100,
			jumpSpeed: 5,
			moveSpeed: 10,
			face: "right",
			dx: 0,
			character: img2,
			jetPack: document.getElementById("jetPackImage"),
			dy: 0,
			x: 0,
			y: 470,
			xInit: 0,
			yInit: 470,
			width: 30,
			height: 50,
			draw: function(){
				this.y += this.dy;
				if(this.y <= this.yInit - this.jumpHeight){
					this.dy = -this.dy;
				}else if(this.y + this.width >= ground.y){
					this.y = ground.y - this.width;
					this.dy = 0;
				}
				if(this.dx > 0){
					this.face = "right";
					this.character = img2;	
				}else if(this.dx < 0){
					this.face = "left";
					this.character = img;
				}
				canvas2D.drawImage(this.character, this.x, this.y, this.width, this.height);
	
	
			},
			move: function(){
				if(keydown.left && !paused){
					player.dx = -player.moveSpeed;
				
				}else if(keydown.right && !paused){
					this.dx = +player.moveSpeed;
				}else
					this.dx = 0;
	
	
				if((keydown.space  || keydown.return) && bulletsLeft != 0 && !paused){
					shot.pause();
					shot.currentTime = 0;
					shot.play();
	
					if(this.face == "right"){
						bullets.push(new Bullet(this.x + this.width/2 - 20, this.y + this.height/2, this.face));
						bulletsLeft--;
					}else{
						bullets.push(new Bullet(this.x - 20, this.y + this.height/2, this.face));
						bulletsLeft--;
	
					}
					
					if(bullets.length >= 50){
						//screamSound.play();
						bullets = [];
					}
				}
				
				player.x += player.dx;
				
				if(keydown.up && fuelLeft != 0 && !paused){
					bubbles.push(new Bubble(this.x + this.width/2, this.y + this.height, 1));
					fuelLeft--;
					if(bubbles.length >= 50){
					//screamSound.play();
					bubbles = [];
				}
	
					this.jump();
	
				}
				
				
				this.x = this.x.clamp(0, CANVAS_WIDTH - player.width);
				if(this.y <= 0){
					this.dy = -this.dy
					this.y = this.dy + 10;
				}
			},
			jump: function(){
				//canvas2D.drawImage(this.jetPack, this.x, this.y, this.width/2, this.height);			
	
				this.yInit = this.y;
				jumpSound.pause();
				jumpSound.currentTime = 0;
				this.y = this.y-10;
				jumpSound.play();
				this.dy = -2*this.jumpSpeed;
			}
			
		};
	
		function Bubble(x, y, dy){
			this.x = x;
			this.y = y;
			this.yInit = y;
			this.dy = dy;
			this.image = document.getElementById("bubbleImage");
			this.draw = function(){
				canvas2D.drawImage(this.image, this.x, this.y, 15, 15);	
				
			};
			this.move = function(){
				this.y = this.y + this.dy;
			};
		}
		
		
		function Enemy(x, y, dx, type){
			this.type = type;			
			this.x = x;
			this.yInit = (y > 100)? y : 100;
			this.y = this.yInit;
			this.dx = dx;
			this.width = 50;
			this.height = 50;
			this.health = 1;
			this.image = document.getElementById("enemy1");		
			
			if(this.type == 1){
				this.image = document.getElementById("enemy2");
				this.health = 2;
			}else if(this.type == 2){
				this.image = document.getElementById("enemy2");
				this.health = 2;

			}else if(this.type == 3){
				this.image = document.getElementById("enemy3");
				this.health = 3;

			}else if(this.type == 4){
				this.image = document.getElementById("enemy4");
				this.health = 4;

			}

			
			//console.log("Type: " + this.type + " Health: " + this.health);


			this.draw = function(){
				canvas2D.drawImage(this.image, this.x, this.y, this.width, this.height);	
				
			};
			this.move = function(){
				this.x = this.x + this.dx;
				this.y = 100*Math.sin((this.x-800)/100) + this.yInit
			};
			
		}
		
		
		function Bonus(x, y, dy, type){
			this.type = 4 - type;
			this.x = x;
			this.y = y;
			this.dy = dy;
			this.width = 30;
			this.height = 30;
			this.image = document.getElementById("bonus1");
			
			if(this.type == 1){
				this.image = document.getElementById("bonus1");
			}else if(this.type == 2){
				this.image = document.getElementById("bonus2");

			}else if(this.type == 3){
				this.image = document.getElementById("bonus3");
			}

			this.draw = function(){
				canvas2D.drawImage(this.image, this.x, this.y, this.width, this.height);	
				
			};
			this.move = function(){
				this.y += this.dy;
			};
			
		}
		
}