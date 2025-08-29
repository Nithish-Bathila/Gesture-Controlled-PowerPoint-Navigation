import cv2
import mediapipe as mp
import time

mp_hands = mp.solutions.hands
hands = mp_hands.Hands(max_num_hands=1, min_detection_confidence=0.7)
mp_draw = mp.solutions.drawing_utils

class GestureDetector:
    def __init__(self):
        self.prev_gesture = None
        self.gesture_start_time = None
        self.hold_time_required = 2  # seconds

    def detect_gesture(self, frame):
        gesture = None

        rgb_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        results = hands.process(rgb_frame)

        if results.multi_hand_landmarks:
            for hand_landmarks in results.multi_hand_landmarks:
                landmarks = hand_landmarks.landmark

                fingers_open = self.count_fingers_open(landmarks)

                # Determine gesture based on fingers open
                if fingers_open == [1, 1, 1, 1, 1]:
                    gesture = 'start_slideshow'  # palm
                elif fingers_open == [0, 0, 0, 0, 0]:
                    gesture = 'end_slideshow'    # fist
                elif fingers_open == [0, 1, 0, 0, 0]:
                    gesture = 'next_slide'       # number 1
                elif fingers_open == [0, 1, 1, 0, 0]:
                    gesture = 'prev_slide'       # number 2

                mp_draw.draw_landmarks(frame, hand_landmarks, mp_hands.HAND_CONNECTIONS)

        confirmed_gesture = self.handle_hold(gesture)
        return confirmed_gesture

    @staticmethod
    def count_fingers_open(landmarks):
        fingers = []

        # Thumb: compare tip and IP joint x for left/right hand
        if landmarks[4].x < landmarks[3].x:
            fingers.append(1)
        else:
            fingers.append(0)

        # Other fingers: tip y < pip y means open
        tips_ids = [8, 12, 16, 20]
        pip_ids = [6, 10, 14, 18]

        for tip_id, pip_id in zip(tips_ids, pip_ids):
            if landmarks[tip_id].y < landmarks[pip_id].y:
                fingers.append(1)
            else:
                fingers.append(0)

        return fingers

    def handle_hold(self, gesture):
        current_time = time.time()

        if gesture != self.prev_gesture:
            self.prev_gesture = gesture
            self.gesture_start_time = current_time
            return None

        if gesture is None:
            self.gesture_start_time = None
            return None

        if self.gesture_start_time and (current_time - self.gesture_start_time >= self.hold_time_required):
            self.gesture_start_time = None
            return gesture

        return None
