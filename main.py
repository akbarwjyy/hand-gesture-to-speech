import cv2
import mediapipe as mp
import numpy as np
import threading
import time
from collections import deque

# Import Windows TTS
try:
    import win32com.client
    WINDOWS_TTS_AVAILABLE = True
except ImportError:
    WINDOWS_TTS_AVAILABLE = False
    print("WARNING: Windows TTS not available!")

class HandGestureSpeech:
    def __init__(self):
        # Initialize MediaPipe Hands
        self.mp_hands = mp.solutions.hands
        self.hands = self.mp_hands.Hands(
            static_image_mode=False,
            max_num_hands=1,
            min_detection_confidence=0.7,
            min_tracking_confidence=0.5
        )
        self.mp_draw = mp.solutions.drawing_utils
        
        # Initialize Windows TTS (SAPI)
        self.tts_engine = None
        if WINDOWS_TTS_AVAILABLE:
            try:
                print("Initializing Windows SAPI TTS...")
                self.tts_engine = win32com.client.Dispatch("SAPI.SpVoice")
                
                # Get available voices
                voices = self.tts_engine.GetVoices()
                print(f"Found {voices.Count} voices:")
                for i in range(voices.Count):
                    voice = voices.Item(i)
                    print(f"  {i}: {voice.GetDescription()}")
                
                # Set to female voice (Zira) - comment out the next 4 lines to use male voice instead
                if voices.Count > 1:
                    self.tts_engine.Voice = voices.Item(1)
                    print(f"Using female voice: {voices.Item(1).GetDescription()}")
                else:
                    # Fallback to first voice if only one available
                    self.tts_engine.Voice = voices.Item(0)
                    print(f"Using voice: {voices.Item(0).GetDescription()}")
                
                # Uncomment the next 3 lines to use male voice (David) instead:
                # if voices.Count > 0:
                #     self.tts_engine.Voice = voices.Item(0)
                #     print(f"Using male voice: {voices.Item(0).GetDescription()}")
                    
                print("Windows SAPI TTS initialized successfully!")
            except Exception as e:
                print(f"Error initializing Windows SAPI: {e}")
                self.tts_engine = None
        else:
            print("Windows TTS not available on this system")
        
        # Gesture state management
        self.last_gesture = None
        self.last_spoken_time = 0
        self.cooldown = 0.8  # seconds
        
        # For gesture stability: require 3 consecutive frames
        self.gesture_buffer = deque(maxlen=3)
        
        # Gesture to speech mapping
        self.gesture_speech_map = {
            "Open Palm": "Halo",
            "Fist": "Saya",
            "Index Only": "Akbar Wijaya",
            "Peace": "Mahasiswa Informatika",
            "Thumbs Up": "Terima kasih"
        }
        
        # For FPS calculation
        self.prev_time = time.time()
        self.fps = 0

    def speak(self, text):
        """Speak text using Windows SAPI"""
        if self.tts_engine is None:
            print(f"Cannot speak: TTS engine not initialized")
            return False
            
        try:
            print(f"Speaking: '{text}'")
            self.tts_engine.Speak(text)
            print("Speech completed successfully")
            return True
        except Exception as e:
            print(f"Error speaking text: {e}")
            return False

    def change_voice(self, voice_index):
        """Change TTS voice"""
        if self.tts_engine is None:
            print("TTS engine not available")
            return False
            
        try:
            voices = self.tts_engine.GetVoices()
            if 0 <= voice_index < voices.Count:
                self.tts_engine.Voice = voices.Item(voice_index)
                print(f"Voice changed to: {voices.Item(voice_index).GetDescription()}")
                return True
            else:
                print(f"Invalid voice index: {voice_index}")
                return False
        except Exception as e:
            print(f"Error changing voice: {e}")
            return False

    def is_finger_extended(self, landmarks, handedness, finger_name):
        """
        Determine if a finger is extended based on MediaPipe landmarks.
        Rules:
        - For index, middle, ring, pinky: tip.y < pip.y (y increases downward)
        - For thumb: depends on handedness
          Right hand: thumb tip.x < ip.x (thumb extends to the left)
          Left hand: thumb tip.x > ip.x (thumb extends to the right)
        """
        # Landmark indices
        finger_tips = {'thumb': 4, 'index': 8, 'middle': 12, 'ring': 16, 'pinky': 20}
        finger_pips = {'thumb': 3, 'index': 6, 'middle': 10, 'ring': 14, 'pinky': 18}
        finger_ips = {'thumb': 2}  # For thumb IP joint
        
        tip_idx = finger_tips[finger_name]
        if finger_name == 'thumb':
            ip_idx = finger_ips['thumb']
            tip = landmarks[tip_idx]
            ip = landmarks[ip_idx]
            
            # Check handedness
            is_right_hand = handedness.classification[0].label == 'Right'
            
            if is_right_hand:
                # Right hand: thumb extended when tip is to the left of IP
                return tip.x < ip.x - 0.02  # Add small margin for stability
            else:
                # Left hand: thumb extended when tip is to the right of IP
                return tip.x > ip.x + 0.02
        else:
            # For other fingers: extended when tip is above PIP (smaller y value)
            pip_idx = finger_pips[finger_name]
            tip = landmarks[tip_idx]
            pip = landmarks[pip_idx]
            return tip.y < pip.y - 0.02  # Add margin for stability

    def classify_gesture(self, landmarks, handedness):
        """Classify hand gesture based on finger states"""
        # Get finger states
        fingers = {}
        finger_names = ['thumb', 'index', 'middle', 'ring', 'pinky']
        
        for finger in finger_names:
            fingers[finger] = self.is_finger_extended(landmarks, handedness, finger)
        
        # Classify based on rules
        if all(fingers.values()):
            return "Open Palm"
        elif not any(fingers.values()):
            return "Fist"
        elif (fingers['index'] and 
              not fingers['middle'] and 
              not fingers['ring'] and 
              not fingers['pinky']):
            return "Index Only"
        elif (fingers['index'] and 
              fingers['middle'] and 
              not fingers['ring'] and 
              not fingers['pinky']):
            return "Peace"
        elif (fingers['thumb'] and 
              not fingers['index'] and 
              not fingers['middle'] and 
              not fingers['ring'] and 
              not fingers['pinky']):
            # Optional: Check if thumb is pointing upward
            # Calculate palm center (average of MCP joints)
            mcp_indices = [5, 9, 13, 17]  # MCP joints of index, middle, ring, pinky
            palm_y = sum(landmarks[i].y for i in mcp_indices) / len(mcp_indices)
            thumb_tip_y = landmarks[4].y
            
            # Thumb should be higher (smaller y) than palm center
            if thumb_tip_y < palm_y - 0.05:
                return "Thumbs Up"
            else:
                return None  # Not a proper thumbs up
        else:
            return None

    def process_frame(self, frame):
        """Process a single frame and return annotated frame with gesture"""
        # Convert to RGB for MediaPipe
        rgb_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        results = self.hands.process(rgb_frame)
        
        current_gesture = None
        
        if results.multi_hand_landmarks and results.multi_handedness:
            for hand_landmarks, handedness in zip(results.multi_hand_landmarks, results.multi_handedness):
                # Draw hand landmarks
                self.mp_draw.draw_landmarks(
                    frame, hand_landmarks, self.mp_hands.HAND_CONNECTIONS,
                    self.mp_draw.DrawingSpec(color=(0, 255, 0), thickness=2, circle_radius=3),
                    self.mp_draw.DrawingSpec(color=(255, 0, 0), thickness=2)
                )
                
                # Get gesture classification
                gesture = self.classify_gesture(hand_landmarks.landmark, handedness)
                current_gesture = gesture
                
                # Draw bounding box
                h, w, _ = frame.shape
                x_coords = [lm.x * w for lm in hand_landmarks.landmark]
                y_coords = [lm.y * h for lm in hand_landmarks.landmark]
                x_min, x_max = int(min(x_coords)), int(max(x_coords))
                y_min, y_max = int(min(y_coords)), int(max(y_coords))
                
                # Add padding
                padding = 20
                x_min = max(0, x_min - padding)
                y_min = max(0, y_min - padding)
                x_max = min(w, x_max + padding)
                y_max = min(h, y_max + padding)
                
                cv2.rectangle(frame, (x_min, y_min), (x_max, y_max), (0, 255, 0), 2)
                
                # Display speech text instead of gesture name
                if gesture:
                    speech_text = self.gesture_speech_map.get(gesture, "")
                    cv2.putText(frame, speech_text, (x_min, y_min - 10), 
                               cv2.FONT_HERSHEY_SIMPLEX, 0.8, (0, 255, 0), 2)
        
        # Add gesture to buffer for stability
        self.gesture_buffer.append(current_gesture)
        
        # Check if we have consistent gesture (3 frames in a row)
        stable_gesture = None
        if len(self.gesture_buffer) == 3:
            if all(g == self.gesture_buffer[0] for g in self.gesture_buffer) and self.gesture_buffer[0] is not None:
                stable_gesture = self.gesture_buffer[0]
        
        # Handle TTS with cooldown
        current_time = time.time()
        if (stable_gesture and 
            stable_gesture != self.last_gesture and 
            current_time - self.last_spoken_time > self.cooldown):
            speech_text = self.gesture_speech_map.get(stable_gesture, "")
            if speech_text:
                # Create a new thread to handle speech so it doesn't block the UI
                speech_thread = threading.Thread(target=self.speak, args=(speech_text,))
                speech_thread.daemon = True
                speech_thread.start()
                
                self.last_spoken_time = current_time
                self.last_gesture = stable_gesture
        
        # Calculate and display FPS
        current_time = time.time()
        self.fps = 1 / (current_time - self.prev_time) if current_time > self.prev_time else 0
        self.prev_time = current_time
        
        cv2.putText(frame, f"FPS: {int(self.fps)}", (10, 30), 
                   cv2.FONT_HERSHEY_SIMPLEX, 0.8, (255, 0, 0), 2)
        
        if stable_gesture:
            speech_text = self.gesture_speech_map.get(stable_gesture, "")
            cv2.putText(frame, f"Detected: {speech_text}", (10, 60), 
                       cv2.FONT_HERSHEY_SIMPLEX, 0.8, (0, 255, 0), 2)
        
        return frame

    def run(self):
        """Main loop"""
        cap = cv2.VideoCapture(0)
        
        if not cap.isOpened():
            print("Error: Could not open webcam")
            return
        
        print("Hand Gesture Speech Recognition Started")
        print("Show gestures:")
        print(" - Open Palm → Halo")
        print(" - Fist → Saya")
        print(" - Index Only → Akbar Wijaya")
        print(" - Peace (V) → Mahasiswa Informatika")
        print(" - Thumbs Up → Terima kasih")
        print("Press ESC to exit")
        
        while True:
            ret, frame = cap.read()
            if not ret:
                break
                
            # Flip the frame horizontally to remove mirror effect
            frame = cv2.flip(frame, 1)
            
            # Process frame
            processed_frame = self.process_frame(frame)
            
            # Display frame
            cv2.imshow('Hand Gesture Speech Recognition', processed_frame)
            
            # Exit on ESC
            key = cv2.waitKey(1) & 0xFF
            if key == 27:  # ESC key
                break
        
        # Cleanup
        cap.release()
        cv2.destroyAllWindows()
        
        print("Program terminated")


if __name__ == "__main__":
    print("=== Hand Gesture to Speech System ===")
    
    # Test Windows SAPI directly
    if WINDOWS_TTS_AVAILABLE:
        try:
            print("Testing Windows SAPI...")
            test_tts = win32com.client.Dispatch("SAPI.SpVoice")
            test_tts.Speak("Test suara sistem")
            print("Windows SAPI test completed successfully!")
        except Exception as e:
            print(f"Windows SAPI test failed: {e}")
    else:
        print("Windows TTS not available!")
        exit(1)
    
    print("\n=== Starting Application ===")
    
    # Create application
    app = HandGestureSpeech()
    
    # Test application TTS
    print("Testing application TTS...")
    app.speak("Sistem Pengenalan Gestur Tangan Siap Digunakan")
    
    # Ask for voice preference
    if app.tts_engine:
        try:
            voices = app.tts_engine.GetVoices()
            if voices.Count > 1:
                print(f"\nAvailable voices:")
                for i in range(voices.Count):
                    voice = voices.Item(i)
                    desc = voice.GetDescription()
                    print(f"  {i}: {desc}")
                
                print("Using default voice (0: David - Male). To change, modify the code.")
                # Uncomment the next lines if you want to choose voice interactively:
                # choice = input(f"Choose voice (0-{voices.Count-1}, or press Enter for default): ").strip()
                # if choice.isdigit():
                #     voice_idx = int(choice)
                #     if app.change_voice(voice_idx):
                #         app.speak("Suara telah diubah")
        except Exception as e:
            print(f"Error in voice selection: {e}")
    
    # Run the main application
    app.run()