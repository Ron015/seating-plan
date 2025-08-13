import random
import logging
from typing import List, Dict, Tuple, Optional

class SeatingAlgorithm:
    def __init__(self):
        self.logger = logging.getLogger(__name__)
    
    def generate_seating_plan(self, students: List[Dict], rooms: List[Dict], 
                            boy_girl_pairing: bool = False,
                            gender_separation: bool = False,
                            random_assignment: bool = True) -> Tuple[Dict, List]:
        """
        Generate seating plan with anti-cheating rules.
        
        Returns:
            Tuple of (seating_plan_dict, unassigned_students_list)
        """
        try:
            # Prepare students
            available_students = students.copy()
            if random_assignment:
                random.shuffle(available_students)
            
            seating_plan = {}
            unassigned_students = []
            
            # Process each room
            for room in rooms:
                room_id = room['room_id']
                rows = room['rows']
                columns = room['columns']
                extra_desks = room.get('extra_desks', 0)
                
                # Initialize room seating
                room_seating = {}
                
                # Add grid desks
                for row in range(1, rows + 1):
                    for col in range(1, columns + 1):
                        desk_id = f"R{row}C{col}"
                        room_seating[desk_id] = {
                            'row': row,
                            'col': col,
                            'seat_a': None,
                            'seat_b': None,
                            'is_extra': False
                        }
                
                # Add extra desks
                for extra in range(1, extra_desks + 1):
                    desk_id = f"E{extra}"
                    room_seating[desk_id] = {
                        'row': rows + 1,  # Place extra desks in a virtual row
                        'col': extra,
                        'seat_a': None,
                        'seat_b': None,
                        'is_extra': True
                    }
                
                # Assign students to this room
                available_students = self._assign_students_to_room(
                    available_students, room_seating, boy_girl_pairing, gender_separation
                )
                
                seating_plan[room_id] = room_seating
            
            # Any remaining students are unassigned
            unassigned_students = available_students
            
            return seating_plan, unassigned_students
            
        except Exception as e:
            self.logger.error(f"Error in generate_seating_plan: {str(e)}")
            return {}, students
    
    def _assign_students_to_room(self, students: List[Dict], room_seating: Dict, 
                               boy_girl_pairing: bool, gender_separation: bool) -> List[Dict]:
        """
        Assign students to a specific room following anti-cheating rules.
        """
        remaining_students = students.copy()
        
        # Get all desk positions
        desk_positions = list(room_seating.keys())
        if not desk_positions:
            return remaining_students
        
        for desk_id in desk_positions:
            if not remaining_students:
                break
                
            desk = room_seating[desk_id]
            
            # Try to assign seat A
            if desk['seat_a'] is None:
                student_a = self._find_suitable_student(
                    remaining_students, desk, room_seating, 'seat_a', boy_girl_pairing, gender_separation
                )
                if student_a:
                    desk['seat_a'] = student_a
                    remaining_students.remove(student_a)
            
            # Try to assign seat B
            if desk['seat_b'] is None and remaining_students:
                student_b = self._find_suitable_student(
                    remaining_students, desk, room_seating, 'seat_b', boy_girl_pairing, gender_separation
                )
                if student_b:
                    desk['seat_b'] = student_b
                    remaining_students.remove(student_b)
        
        return remaining_students
    
    def _find_suitable_student(self, students: List[Dict], desk: Dict, 
                             room_seating: Dict, seat_position: str, 
                             boy_girl_pairing: bool, gender_separation: bool) -> Optional[Dict]:
        """
        Find a student that can be placed at the given desk and seat position
        without violating anti-cheating rules.
        """
        for student in students:
            if self._can_place_student(student, desk, room_seating, seat_position, boy_girl_pairing, gender_separation):
                return student
        return None
    
    def _can_place_student(self, student: Dict, desk: Dict, 
                         room_seating: Dict, seat_position: str, 
                         boy_girl_pairing: bool, gender_separation: bool) -> bool:
        """
        Check if a student can be placed without violating rules.
        """
        student_class = student['class']
        student_gender = student['gender']
        
        # Same desk check
        other_seat = 'seat_a' if seat_position == 'seat_b' else 'seat_b'
        other_student = desk.get(other_seat)
        if other_student:
            if other_student['class'] == student_class:
                return False
            if gender_separation and other_student['gender'] != student_gender:
                return False
            if boy_girl_pairing and other_student['gender'] == student_gender:
                return False
        
        # Adjacent desk same class check
        if self._has_same_class_adjacent(student_class, desk, room_seating):
            return False
        
        # Temporary place student for deeper check
        original = desk[seat_position]
        desk[seat_position] = student
        violations = self._validate_partial(room_seating)
        desk[seat_position] = original
        
        if violations:
            return False
        
        return True
    

    def _validate_partial(self, room_seating: Dict) -> bool:
        """Check if current partial seating already violates same class rule."""
        for desk_id, desk in room_seating.items():
            seat_a = desk.get('seat_a')
            seat_b = desk.get('seat_b')
            if seat_a and seat_b and seat_a['class'] == seat_b['class']:
                return True
            # Adjacent check
            for seat_pos in ['seat_a', 'seat_b']:
                student = desk.get(seat_pos)
                if student and self._has_same_class_adjacent(student['class'], desk, room_seating):
                    return True
        return False
        
    def _has_same_class_adjacent(self, student_class: str, target_desk: Dict, 
                               room_seating: Dict) -> bool:
        """
        Check if any adjacent desk has a student from the same class.
        """
        row = target_desk['row']
        col = target_desk['col']
        
        # Check all four adjacent positions
        adjacent_positions = [
            (row - 1, col),  # Above
            (row + 1, col),  # Below
            (row, col - 1),  # Left
            (row, col + 1),  # Right
        ]
        
        for adj_row, adj_col in adjacent_positions:
            adj_desk_id = f"R{adj_row}C{adj_col}"
            
            if adj_desk_id in room_seating:
                adj_desk = room_seating[adj_desk_id]
                
                # Check both seats in adjacent desk
                for seat in ['seat_a', 'seat_b']:
                    adj_student = adj_desk.get(seat)
                    if adj_student and adj_student['class'] == student_class:
                        return True
        
        return False
    
    def validate_seating_plan(self, seating_plan: Dict) -> List[str]:
        """
        Validate that the seating plan follows all anti-cheating rules.
        Returns a list of violations found.
        """
        violations = []
        
        for room_id, room_seating in seating_plan.items():
            for desk_id, desk in room_seating.items():
                # Check same desk rule
                seat_a = desk.get('seat_a')
                seat_b = desk.get('seat_b')
                
                if seat_a and seat_b and seat_a['class'] == seat_b['class']:
                    violations.append(
                        f"Room {room_id}, Desk {desk_id}: Same class students on same desk"
                    )
                
                # Check adjacent desk rules
                for seat_pos in ['seat_a', 'seat_b']:
                    student = desk.get(seat_pos)
                    if student and self._has_same_class_adjacent(student['class'], desk, room_seating):
                        violations.append(
                            f"Room {room_id}, Desk {desk_id}, {seat_pos}: Adjacent same class violation"
                        )
        
        return violations


# Global wrapper functions for convenience
def generate_seating_plan(students: List[Dict], rooms: List[Dict], 
                         boy_girl_pairing: bool = False,
                         gender_separation: bool = False,
                         random_assignment: bool = True) -> Tuple[Dict, List]:
    """
    Convenience wrapper for the SeatingAlgorithm class.
    """
    algorithm = SeatingAlgorithm()
    return algorithm.generate_seating_plan(
        students, rooms, boy_girl_pairing, gender_separation, random_assignment
    )


# Global instance for import
seating_algorithm = SeatingAlgorithm()
